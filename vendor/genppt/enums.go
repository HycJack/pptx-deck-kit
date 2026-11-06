// Package genppt 提供纯Go语言的PowerPoint演示文稿生成功能
// 基于OOXML (ECMA-376)标准，生成兼容Microsoft PowerPoint、Apple Keynote、
// LibreOffice Impress的.pptx文件
package genppt

// SlideLayout 定义幻灯片布局类型
type SlideLayout string

const (
	// LayoutBlank 空白布局
	LayoutBlank SlideLayout = "blank"
	// LayoutTitle 标题布局
	LayoutTitle SlideLayout = "title"
	// LayoutTitleContent 标题和内容布局
	LayoutTitleContent SlideLayout = "titleContent"
	// LayoutTwoContent 两栏内容布局
	LayoutTwoContent SlideLayout = "twoContent"
)

// ShapeType 定义形状类型
type ShapeType string

const (
	// ShapeRect 矩形
	ShapeRect ShapeType = "rect"
	// ShapeRoundRect 圆角矩形
	ShapeRoundRect ShapeType = "roundRect"
	// ShapeEllipse 椭圆
	ShapeEllipse ShapeType = "ellipse"
	// ShapeTriangle 三角形
	ShapeTriangle ShapeType = "triangle"
	// ShapeDiamond 菱形
	ShapeDiamond ShapeType = "diamond"
	// ShapeArrowRight 右箭头
	ShapeArrowRight ShapeType = "rightArrow"
	// ShapeArrowLeft 左箭头
	ShapeArrowLeft ShapeType = "leftArrow"
	// ShapeArrowUp 上箭头
	ShapeArrowUp ShapeType = "upArrow"
	// ShapeArrowDown 下箭头
	ShapeArrowDown ShapeType = "downArrow"
	// ShapeStar5 五角星
	ShapeStar5 ShapeType = "star5"
	// ShapeHeart 心形
	ShapeHeart ShapeType = "heart"
	// ShapeLine 直线
	ShapeLine ShapeType = "line"
)

// Align 定义水平对齐方式
type Align string

const (
	// AlignLeft 左对齐
	AlignLeft Align = "l"
	// AlignCenter 居中对齐
	AlignCenter Align = "ctr"
	// AlignRight 右对齐
	AlignRight Align = "r"
	// AlignJustify 两端对齐
	AlignJustify Align = "just"
)

// VerticalAlign 定义垂直对齐方式
type VerticalAlign string

const (
	// VAlignTop 顶部对齐
	VAlignTop VerticalAlign = "t"
	// VAlignMiddle 垂直居中
	VAlignMiddle VerticalAlign = "ctr"
	// VAlignBottom 底部对齐
	VAlignBottom VerticalAlign = "b"
)

// BorderStyle 定义边框样式
type BorderStyle string

const (
	// BorderSolid 实线
	BorderSolid BorderStyle = "solid"
	// BorderDash 虚线
	BorderDash BorderStyle = "dash"
	// BorderDot 点线
	BorderDot BorderStyle = "dot"
	// BorderNone 无边框
	BorderNone BorderStyle = "none"
)

// OutputType 定义输出类型
type OutputType string

const (
	// OutputFile 输出到文件
	OutputFile OutputType = "file"
	// OutputBytes 输出为字节数组
	OutputBytes OutputType = "bytes"
	// OutputWriter 输出到io.Writer
	OutputWriter OutputType = "writer"
)

// AnimEffect 定义入场动画效果
type AnimEffect string

const (
	// AnimAppear 出现（无过渡）
	AnimAppear AnimEffect = "appear"
	// AnimFade 淡入
	AnimFade AnimEffect = "fade"
	// AnimWipe 擦除（自下往上）
	AnimWipe AnimEffect = "wipe"
	// AnimFly 飞入（自左）
	AnimFly AnimEffect = "fly"
)

// preset 返回 OOXML 的 presetID 与 presetSubtype。
// 这三个编号是 PowerPoint 自己的编号，实测回读一致。
func (a AnimEffect) preset() (int, int) {
	switch a {
	case AnimAppear:
		return 1, 0
	case AnimWipe:
		return 22, 4
	case AnimFly:
		return 2, 4
	default:
		return 10, 0
	}
}

// TransitionKind 定义幻灯片切换效果。
type TransitionKind string

// 这里只收录 PowerPoint 16.0 能完整往返的写法：实测把
// <p:transition><p14:morph option="byObject"/></p:transition> 写进文件后，
// PowerPoint 打开再另存会把节点丢掉，所以 morph 没有收录。
const (
	// TransitionNone 无切换
	TransitionNone TransitionKind = "none"
	// TransitionFade 淡出淡入
	TransitionFade TransitionKind = "fade"
	// TransitionWipe 自下擦除
	TransitionWipe TransitionKind = "wipe"
	// TransitionPush 自下推入
	TransitionPush TransitionKind = "push"
	// TransitionCut 直接切换
	TransitionCut TransitionKind = "cut"
	// TransitionCircle 圆形展开
	TransitionCircle TransitionKind = "circle"
	// TransitionDiamond 菱形展开
	TransitionDiamond TransitionKind = "diamond"
	// TransitionWedge 楔形展开
	TransitionWedge TransitionKind = "wedge"
	// TransitionPlus 十字展开
	TransitionPlus TransitionKind = "plus"
	// TransitionWheel 单轮辐旋转展开
	TransitionWheel TransitionKind = "wheel"
	// TransitionNewsflash 快闪
	TransitionNewsflash TransitionKind = "newsflash"
)

// transitionTags 切换效果对应的 OOXML 片段
var transitionTags = map[TransitionKind]string{
	TransitionFade:      `<p:fade/>`,
	TransitionWipe:      `<p:wipe dir="u"/>`,
	TransitionPush:      `<p:push dir="u"/>`,
	TransitionCut:       `<p:cut/>`,
	TransitionCircle:    `<p:circle/>`,
	TransitionDiamond:   `<p:diamond/>`,
	TransitionWedge:     `<p:wedge/>`,
	TransitionPlus:      `<p:plus/>`,
	TransitionWheel:     `<p:wheel spokes="1"/>`,
	TransitionNewsflash: `<p:newsflash/>`,
}
