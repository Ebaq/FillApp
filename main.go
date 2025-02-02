package main

import (
	"embed"
	"github.com/wailsapp/wails/v2"
	"github.com/wailsapp/wails/v2/pkg/options"
	"github.com/wailsapp/wails/v2/pkg/options/assetserver"
)

//go:embed all:frontend/dist
var assets embed.FS

func main() {
	// Create an instance of the app structure
	app := NewApp()

	//mainMenu := menu.NewMenu()
	//
	//guideMenu := mainMenu.AddSubmenu("Инструкция")
	//guideMenu.AddText("Открыть инструкцию", keys.CmdOrCtrl("p"), func(_ *menu.CallbackData) {
	//	println("Инструкция")
	//	_ = wails.Run(&options.App{
	//		Title:  "FillApp",
	//		Width:  1440,
	//		Height: 1024,
	//		AssetServer: &assetserver.Options{
	//			Assets: assets,
	//		},
	//		BackgroundColour: &options.RGBA{R: 27, G: 38, B: 54, A: 1},
	//		OnStartup:        app.startup,
	//		Menu:             mainMenu,
	//		Bind: []interface{}{
	//			app,
	//		},
	//	})
	//}

	// Create application with options
	err := wails.Run(&options.App{
		Title:  "FillApp",
		Width:  1440,
		Height: 1024,
		AssetServer: &assetserver.Options{
			Assets: assets,
		},
		BackgroundColour: &options.RGBA{R: 255, G: 255, B: 255, A: 1},
		OnStartup:        app.startup,
		Bind: []interface{}{
			app,
		},
	})

	if err != nil {
		println("Error:", err.Error())
	}
}
