package main

import (
	"context"
	"fmt"

	"github.com/wailsapp/wails/v2/pkg/runtime"
)

// App struct
type App struct {
	ctx context.Context
}

// NewApp creates a new App application struct
func NewApp() *App {
	return &App{}
}

// startup is called when the app starts. The context is saved
// so we can call the runtime methods
func (a *App) startup(ctx context.Context) {
	a.ctx = ctx
}

// Greet returns a greeting for the given name
func (a *App) Greet(name string) string {
	return fmt.Sprintf("Hello %s, It's show time!", name)
}

func (a *App) SelectFiles() ([]string, error) {
	file, err := runtime.OpenMultipleFilesDialog(a.ctx, runtime.OpenDialogOptions{
		Title: "Select Excel File",
		Filters: []runtime.FileFilter{
			{
				DisplayName: "Excel Files",
				Pattern:     "*.xlsx",
			},
		},
	})
	if err != nil {
		return nil, err
	}
	return file, nil
}

func (a *App) ProcessFiles(filePaths []string) ([]string, error) {
	log := []string{}
	for _, path := range filePaths {
		err := Excel2HTML(path)
		if err != nil {
			log = append(log, fmt.Sprintf("failed : %s : %s", path, err.Error()))
			continue
		}
		log = append(log, fmt.Sprintf("success : %s", path))
	}
	fmt.Printf("%+v", filePaths)
	return log, nil
}
