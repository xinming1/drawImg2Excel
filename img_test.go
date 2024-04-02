package main

import (
	"fmt"
	"image"
	"image/color"
	"image/draw"
	"log"
	"os"
	"testing"

	"github.com/xuri/excelize/v2"

	"github.com/disintegration/imaging"
)

func TestImg(t *testing.T) {
	test1()
}

func test1() {
	// 打开原始图片文件
	file, err := os.Open("img.jpg")
	if err != nil {
		log.Fatal(err)
	}
	defer file.Close()

	// 解码图片文件
	img, _, err := image.Decode(file)
	if err != nil {
		log.Fatal(err)
	}

	// 将图片拆分成 gridSize*gridSize 个小格，并填充颜色
	gridSize := 100
	gridWidth := img.Bounds().Dx() / gridSize
	gridHeight := img.Bounds().Dy() / gridSize

	result := imaging.New(gridWidth*gridSize, gridHeight*gridSize, color.RGBA{})

	// 创建一个新的 Excel 文件
	f := excelize.NewFile()

	// 设置单元格宽度和高度为固定长度
	cellWidth := 5
	cellHeight := 30

	for y := 0; y < gridSize; y++ {
		for x := 0; x < gridSize; x++ {
			// 计算当前小格的边界
			gridRect := image.Rect(x*gridWidth, y*gridHeight, (x+1)*gridWidth, (y+1)*gridHeight)

			// 获取当前小格的颜色占比最大的颜色
			maxColor := getMaxColor(img, gridRect)

			// 填充当前小格的颜色
			draw.Draw(result, gridRect, &image.Uniform{C: maxColor}, image.Point{}, draw.Src)

			// 计算当前单元格的位置
			cellName, _ := excelize.CoordinatesToCellName(x+1, y+1)

			hex := colorToHex(maxColor)
			styleId, err := f.NewStyle(
				&excelize.Style{
					Fill: excelize.Fill{
						Type:    "pattern",
						Pattern: 1,
						Color:   []string{hex},
					},
				},
			)
			// 设置单元格宽度和高度
			err = f.SetColWidth("Sheet1", cellName[:1], cellName[:1], float64(cellWidth))
			if err != nil {
				log.Fatal(err)
			}
			err = f.SetRowHeight("Sheet1", y+1, float64(cellHeight))
			if err != nil {
				log.Fatal(err)
			}

			f.SetCellStyle(
				"Sheet1",
				cellName,
				cellName,
				styleId,
			)

		}
	}

	// 保存 Excel 文件
	err = f.SaveAs("output.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("Excel 文件创建完成！")

	// 保存处理后的图片
	err = imaging.Save(result, "output.jpg")
	if err != nil {
		log.Fatal(err)
	}

	fmt.Println("图片处理完成！")
}
