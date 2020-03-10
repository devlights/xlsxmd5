package main

import (
	"context"
	"crypto/md5"
	"flag"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"

	"github.com/devlights/goxcel"
	"golang.org/x/sync/errgroup"
)

type (
	args struct {
		directory string
		pattern   string
		output    string
	}

	md5result struct {
		path     string
		checkSum [md5.Size]byte
		name     string
	}
)

var (
	cmdArgs = args{}
)

const (
	_NumberOf2ndStageGoroutines = 10
)

func main() {
	os.Exit(run())
}

func run() int {

	flag.StringVar(&cmdArgs.directory, "d", ".", "対象ディレクトリ")
	flag.StringVar(&cmdArgs.pattern, "p", "*.*", "対象ファイルパターン")
	flag.StringVar(&cmdArgs.output, "o", "", "出力ファイルパス")
	flag.Parse()

	if cmdArgs.output == "" {
		flag.Usage()
		return 2
	}

	if cmdArgs.directory == "" {
		cmdArgs.directory = "."
	}

	if cmdArgs.pattern == "" {
		cmdArgs.pattern = "*.*"
	}

	var (
		rootCtx           = context.Background()
		errGrp, errGrpCtx = errgroup.WithContext(rootCtx)
	)

	var (
		filePathCh = make(chan string)
		md5Ch      = make(chan md5result)
	)

	// 1st stage
	start1stStage(errGrp, errGrpCtx, filePathCh)

	// 2nd stage
	start2ndStage(errGrp, errGrpCtx, filePathCh, md5Ch)

	// 3rd stage
	start3rdStage(errGrp, md5Ch)

	// final stage
	execFinalStage(md5Ch)

	if err := errGrp.Wait(); err != nil {
		fmt.Println(err)
		return 1
	}

	return 0
}

func start1stStage(errGrp *errgroup.Group, ctx context.Context, filePathCh chan<- string) {

	errGrp.Go(func() error {
		defer close(filePathCh)
		return filepath.Walk(cmdArgs.directory, func(path string, info os.FileInfo, err error) error {
			if err != nil {
				return err
			}

			if info.IsDir() {
				return nil
			}

			match, _ := regexp.Match(cmdArgs.pattern, []byte(info.Name()))
			if match {
				filePathCh <- path
			}

			select {
			case <-ctx.Done():
				return ctx.Err()
			default:
				return nil
			}
		})
	})
}

func start2ndStage(errGrp *errgroup.Group, ctx context.Context, filePathCh <-chan string, md5Ch chan<- md5result) {

	for i := 0; i < _NumberOf2ndStageGoroutines; i++ {
		goroutineIndex := i + 1
		errGrp.Go(func() error {
			var (
				name  = fmt.Sprintf("goroutine-%02d", goroutineIndex)
				count = 0
			)

			for p := range filePathCh {
				data, err := ioutil.ReadFile(p)
				if err != nil {
					return err
				}

				checksum := md5.Sum(data)
				result := md5result{
					path:     p,
					checkSum: checksum,
					name:     name,
				}

				select {
				case <-ctx.Done():
					return ctx.Err()
				case md5Ch <- result:
					count++
				}
			}

			return nil
		})
	}
}

func start3rdStage(errGrp *errgroup.Group, md5Ch chan md5result) {
	go func() {
		_ = errGrp.Wait()
		close(md5Ch)
	}()
}

func execFinalStage(md5Ch <-chan md5result) {

	quitGoxcelFn, _ := goxcel.InitGoxcel()
	defer quitGoxcelFn()

	g, gr, _ := goxcel.NewGoxcel()
	defer gr()

	_ = g.SetDisplayAlerts(false)
	_ = g.SetVisible(false)

	wbs, _ := g.Workbooks()
	wb, wbr, _ := wbs.Add()
	defer wbr()

	wss, _ := wb.WorkSheets()
	ws, _ := wss.Item(1)

	row := 1
	for r := range md5Ch {
		fileNameCell, _ := ws.Cells(row, 1)
		_ = fileNameCell.SetValue(r.path)

		md5ChecksumCell, _ := ws.Cells(row, 2)
		_ = md5ChecksumCell.SetValue(fmt.Sprintf("%x", r.checkSum))

		row++
	}

	_ = wb.SaveAs(cmdArgs.output)
}
