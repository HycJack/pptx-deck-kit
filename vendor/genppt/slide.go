package genppt

import (
	"os"
)

// AddText 添加文本框
func (s *Slide) AddText(text string, opts TextOptions) *Slide {
	obj := &textObject{
		text:    text,
		options: opts,
	}
	// 设置默认值
	if obj.options.FontFace == "" {
		obj.options.FontFace = getDefaultFontFace()
	}
	if obj.options.FontSize == 0 {
		obj.options.FontSize = getDefaultFontSize()
	}
	if obj.options.FontColor == "" {
		obj.options.FontColor = getDefaultColor()
	}
	if obj.options.Align == "" {
		obj.options.Align = AlignLeft
	}
	if obj.options.VAlign == "" {
		obj.options.VAlign = VAlignTop
	}
	s.objects = append(s.objects, obj)
	return s
}

// AddShape 添加形状
func (s *Slide) AddShape(shapeType ShapeType, opts ShapeOptions) *Slide {
	obj := &shapeObject{
		shapeType: shapeType,
		options:   opts,
	}
	// 设置默认值
	if obj.options.Fill == "" && obj.options.Gradient == nil {
		obj.options.Fill = "4472C4" // Office默认蓝色
	}
	if obj.options.LineStyle == "" {
		obj.options.LineStyle = BorderSolid
	}
	// 无边框时不注入默认描边，否则 PowerPoint 会画出 2F5496 的边框
	if obj.options.LineStyle != BorderNone && obj.options.LineColor != "none" {
		if obj.options.LineWidth == 0 {
			obj.options.LineWidth = 1.0
		}
		if obj.options.LineColor == "" {
			obj.options.LineColor = "2F5496"
		}
	}
	s.objects = append(s.objects, obj)
	return s
}

// AddShapeWithText 添加带文本的形状
func (s *Slide) AddShapeWithText(shapeType ShapeType, text string, opts ShapeOptions) *Slide {
	obj := &shapeObject{
		shapeType: shapeType,
		options:   opts,
		text:      text,
	}
	// 设置默认值
	if obj.options.Fill == "" {
		obj.options.Fill = "4472C4"
	}
	if obj.options.LineWidth == 0 {
		obj.options.LineWidth = 1.0
	}
	if obj.options.LineColor == "" {
		obj.options.LineColor = "2F5496"
	}
	s.objects = append(s.objects, obj)
	return s
}

// AddTable 添加表格
func (s *Slide) AddTable(rows [][]TableCell, opts TableOptions) *Slide {
	obj := &tableObject{
		rows:    rows,
		options: opts,
	}
	// 设置默认值
	if obj.options.FontFace == "" {
		obj.options.FontFace = getDefaultFontFace()
	}
	if obj.options.FontSize == 0 {
		obj.options.FontSize = 14
	}
	if obj.options.FontColor == "" {
		obj.options.FontColor = getDefaultColor()
	}
	if obj.options.Border.Width == 0 {
		obj.options.Border.Width = 1.0
	}
	if obj.options.Border.Color == "" {
		obj.options.Border.Color = "CCCCCC"
	}
	if obj.options.Border.Style == "" {
		obj.options.Border.Style = BorderSolid
	}
	s.objects = append(s.objects, obj)
	return s
}

// AddImage 添加图片
func (s *Slide) AddImage(opts ImageOptions) *Slide {
	var data []byte
	var ext string
	var err error

	if opts.Path != "" {
		// 从文件读取
		data, err = os.ReadFile(opts.Path)
		if err != nil {
			// 图片读取失败，跳过
			return s
		}
		ext = getExtFromPath(opts.Path)
	} else if len(opts.Data) > 0 {
		data = opts.Data
		ext = getImageType(data)
	} else {
		return s
	}

	if ext == "" {
		ext = "png"
	}

	// 生成媒体文件关系ID
	mediaIndex := len(s.presentation.mediaFiles) + 1
	rID := "rId" + itoa(mediaIndex+100) // 使用较大的rId避免冲突
	mediaPath := "ppt/media/image" + itoa(mediaIndex) + "." + ext

	// 添加到演示文稿的媒体文件列表
	s.presentation.mediaFiles = append(s.presentation.mediaFiles, mediaFile{
		path: mediaPath,
		data: data,
		ext:  ext,
		rID:  rID,
	})

	obj := &imageObject{
		options:  opts,
		rID:      rID,
		mediaExt: ext,
	}
	s.objects = append(s.objects, obj)
	return s
}

// SetBackground 设置背景
func (s *Slide) SetBackground(opts BackgroundOptions) *Slide {
	s.background = &opts
	return s
}

// SetNotes 设置备注
func (s *Slide) SetNotes(notes string) *Slide {
	s.notes = notes
	return s
}

// SetLayout 设置幻灯片布局
func (s *Slide) SetLayout(layout SlideLayout) *Slide {
	s.layout = layout
	return s
}

// 图层顺序：objects 切片的顺序就是 spTree 的顺序，即渲染时的叠压顺序。

// SendToBack 把最后添加的对象移到最底层
func (s *Slide) SendToBack() *Slide {
	return s.ZOrder(0)
}

// BringToFront 把最后添加的对象移到最顶层
func (s *Slide) BringToFront() *Slide {
	return s.ZOrder(len(s.objects) - 1)
}

// ZOrder 把最后添加的对象移到指定层（0 为最底层）
func (s *Slide) ZOrder(index int) *Slide {
	n := len(s.objects)
	if n == 0 {
		return s
	}
	if index < 0 {
		index = 0
	}
	if index > n-1 {
		index = n - 1
	}
	obj := s.objects[n-1]
	rest := append([]slideObject{}, s.objects[:n-1]...)
	s.objects = append(rest[:index], append([]slideObject{obj}, rest[index:]...)...)
	return s
}

// SetTransition 设置本页切换效果
func (s *Slide) SetTransition(kind TransitionKind) *Slide {
	s.transition = kind
	return s
}

// AnimateLast 给最后添加的对象加入场动画。
// delay 与 duration 单位为毫秒；duration <= 0 时取 500ms。
func (s *Slide) AnimateLast(effect AnimEffect, delay, duration int) *Slide {
	if len(s.objects) == 0 {
		return s
	}
	if effect == "" {
		effect = AnimFade
	}
	if duration <= 0 {
		duration = 500
	}
	s.animations = append(s.animations, animationEntry{
		objectIndex: len(s.objects) - 1,
		effect:      effect,
		delay:       delay,
		duration:    duration,
	})
	return s
}

// AnimateAll 给当前所有对象按加入顺序错峰加入场动画
func (s *Slide) AnimateAll(effect AnimEffect, stagger, duration int) *Slide {
	for i := range s.objects {
		if duration <= 0 {
			duration = 500
		}
		s.animations = append(s.animations, animationEntry{
			objectIndex: i,
			effect:      effect,
			delay:       i * stagger,
			duration:    duration,
		})
	}
	return s
}

// itoa 简单的整数转字符串（避免导入strconv）
func itoa(n int) string {
	if n == 0 {
		return "0"
	}
	if n < 0 {
		return "-" + itoa(-n)
	}
	var digits []byte
	for n > 0 {
		digits = append([]byte{byte('0' + n%10)}, digits...)
		n /= 10
	}
	return string(digits)
}

// ftoa 简单的浮点数转字符串
func ftoa(f float64) string {
	if f == float64(int64(f)) {
		return itoa(int(f))
	}
	// 保留小数点后两位
	intPart := int64(f)
	fracPart := int64((f - float64(intPart)) * 100)
	if fracPart < 0 {
		fracPart = -fracPart
	}
	result := itoa(int(intPart)) + "."
	if fracPart < 10 {
		result += "0"
	}
	result += itoa(int(fracPart))
	return result
}
