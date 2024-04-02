package main

import (
	"flag"
	"fmt"
	"github.com/disintegration/imaging"
	"github.com/xuri/excelize/v2"
	"image"
	"image/color"
	"image/draw"
	"log"
	"os"
)

func main() {
	var imgPath string
	flag.StringVar(&imgPath, "img", "./test.jpg", "for config path")
	flag.Parse()
	drawExcel(imgPath)
}

func drawExcel(imgPath string) {
	// 打开原始图片文件
	file, err := os.Open(imgPath)
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
	fmt.Println("处理后的图片已保存为 output.xlsx")

	//// 保存处理后的图片
	//err = imaging.Save(result, "output.jpg")
	//if err != nil {
	//	log.Fatal(err)
	//}
	//
	//fmt.Println("图片处理完成！")
}

// 将 color.Color 转换为 16 进制字符串
func colorToHex(c color.Color) string {
	r, g, b, _ := c.RGBA()
	// 将 RGBA 分量值还原为 8 位无符号整数
	r >>= 8
	g >>= 8
	b >>= 8
	rgb := (uint32(r) << 16) | (uint32(g) << 8) | uint32(b)
	return fmt.Sprintf("#%06X", rgb)
}

// 获取指定区域内颜色占比最大的颜色
func getMaxColor(img image.Image, rect image.Rectangle) color.Color {
	colorCount := make(map[color.Color]int)

	// 统计指定区域内每个颜色出现的次数
	for y := rect.Min.Y; y < rect.Max.Y; y++ {
		for x := rect.Min.X; x < rect.Max.X; x++ {
			colorCount[img.At(x, y)]++
		}
	}

	// 找到颜色占比最大的颜色
	maxCount := 0
	var maxColor color.Color

	for c, count := range colorCount {
		if count > maxCount {
			maxCount = count
			maxColor = c
		}
	}

	return maxColor
}
