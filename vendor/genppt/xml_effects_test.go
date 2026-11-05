package genppt

import (
	"strings"
	"testing"
)

// 无边框必须显式写 noFill，且不能注入主题默认描边
func TestShapeBorderNone(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF", LineStyle: BorderNone})
	xml := s.generateSlide()

	if !strings.Contains(xml, `<a:ln><a:noFill/></a:ln>`) {
		t.Error("无边框没有写出 <a:ln><a:noFill/></a:ln>")
	}
	if strings.Contains(xml, "2F5496") {
		t.Error("无边框仍注入了主题默认描边 2F5496")
	}
}

// 显式边框仍按参数写入
func TestShapeBorderSolid(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF",
		LineColor: "D8D0BE", LineWidth: 0.75, LineStyle: BorderSolid})
	xml := s.generateSlide()
	if !strings.Contains(xml, `w="9525"`) || !strings.Contains(xml, "D8D0BE") {
		t.Errorf("边框宽度/颜色没有写入: %s", xml)
	}
}

// 渐变填充：色标、角度、且不覆盖默认纯色
func TestShapeGradient(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRoundRect, ShapeOptions{X: 1, Y: 1, Width: 3, Height: 1,
		Gradient: &GradientFill{
			Stops: []GradientStop{{Pos: 0, Color: "2E4A2A"}, {Pos: 100, Color: "E89CB1"}},
			Angle: 45,
		}})
	xml := s.generateSlide()

	for _, want := range []string{`<a:gradFill`, `pos="0"`, `pos="100000"`, `2E4A2A`, `E89CB1`, `<a:lin ang="2700000"`} {
		if !strings.Contains(xml, want) {
			t.Errorf("渐变缺少 %s", want)
		}
	}
	if strings.Contains(xml, `val="4472C4"`) {
		t.Error("设置了渐变却仍注入默认填充色")
	}
}

// 径向渐变走 path 而不是 lin
func TestShapeGradientRadial(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeEllipse, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 2,
		Gradient: &GradientFill{Radial: true, Stops: []GradientStop{{Pos: 0, Color: "FFFFFF"}, {Pos: 100, Color: "2E4A2A"}}}})
	xml := s.generateSlide()
	if !strings.Contains(xml, `<a:path path="circle">`) {
		t.Error("径向渐变没有写出 a:path")
	}
}

// 阴影参数可达
func TestShapeShadowFX(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF",
		ShadowFX: &ShadowOptions{Blur: 20, Dist: 10, Direction: 90, Color: "1A1A17", Alpha: 18}})
	xml := s.generateSlide()

	// 20pt = 254000 EMU, 10pt = 127000 EMU, 90° = 5400000
	for _, want := range []string{`blurRad="254000"`, `dist="127000"`, `dir="5400000"`, `1A1A17`, `<a:alpha val="18000"/>`} {
		if !strings.Contains(xml, want) {
			t.Errorf("阴影缺少 %s", want)
		}
	}
}

// 图层顺序：SendToBack 后加的对象反而先画
func TestSendToBack(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FF0000"})
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "00FF00"})
	s.SendToBack()
	xml := s.generateSlide()

	green := strings.Index(xml, "00FF00")
	red := strings.Index(xml, "FF0000")
	if green < 0 || red < 0 {
		t.Fatal("找不到两个颜色")
	}
	if green > red {
		t.Error("SendToBack 之后绿色仍然在后面")
	}
}

// 切换必须排在 timing 之前（CT_Slide 的元素次序）
func TestTransitionBeforeTiming(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.SetTransition(TransitionFade)
	s.AddText("hi", TextOptions{X: 1, Y: 1, Width: 3, Height: 1, FontSize: 18})
	s.AnimateLast(AnimFade, 0, 500)
	xml := s.generateSlide()

	tr := strings.Index(xml, "<p:transition")
	tm := strings.Index(xml, "<p:timing>")
	if tr < 0 || tm < 0 {
		t.Fatalf("transition=%d timing=%d", tr, tm)
	}
	if tr > tm {
		t.Error("transition 必须排在 timing 之前")
	}
	if !strings.Contains(xml, `<p:fade/>`) {
		t.Error("淡入切换没有写出 <p:fade/>")
	}
}

// 动画时间轴：四层结构 + 正确的 spid + bldLst
func TestAnimationTiming(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF"})
	s.AnimateLast(AnimFade, 0, 500)
	s.AddShape(ShapeRect, ShapeOptions{X: 4, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF"})
	s.AnimateLast(AnimWipe, 200, 500)
	xml := s.generateSlide()

	// 两个形状的 spid 分别是 2 和 3
	for _, want := range []string{
		`presetID="10"`, `presetID="22"`, `nodeType="clickEffect"`, `nodeType="withEffect"`,
		`<p:cond delay="0"/>`, `<p:cond delay="200"/>`,
		`<p:spTgt spid="2"/>`, `<p:spTgt spid="3"/>`,
		`<p:bldP spid="2" grpId="0"/>`, `<p:bldP spid="3" grpId="0"/>`,
		`<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>`,
		`filter="fade"`, `filter="wipe(up)"`,
	} {
		if !strings.Contains(xml, want) {
			t.Errorf("时间轴缺少 %s", want)
		}
	}

	// par 共 5 层：tmRoot + click组 + 组内容器 + 2 个效果
	if n := strings.Count(xml, "<p:par>"); n != 5 {
		t.Errorf("par 层数应为 5（tmRoot + click组 + 组内容器 + 2 个效果），实际 %d", n)
	}
}

// 飞入用 anim 改 ppt_x，不是简单的 animEffect
func TestAnimationFly(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF"})
	s.AnimateLast(AnimFly, 0, 500)
	xml := s.generateSlide()
	if !strings.Contains(xml, `<p:attrName>ppt_x</p:attrName>`) ||
		!strings.Contains(xml, `val="#ppt_x-1/3"`) {
		t.Error("飞入没有写出 ppt_x 位移关键帧")
	}
}

// 出现只有可见性，没有过渡行为
func TestAnimationAppear(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.AddShape(ShapeRect, ShapeOptions{X: 1, Y: 1, Width: 2, Height: 1, Fill: "FFFFFF"})
	s.AnimateLast(AnimAppear, 0, 500)
	xml := s.generateSlide()
	if !strings.Contains(xml, `presetID="1"`) {
		t.Error("出现效果 presetID 应为 1")
	}
	if strings.Contains(xml, `filter="fade"`) {
		t.Error("出现效果不应带过渡行为")
	}
}

// 渐变背景
func TestGradientBackground(t *testing.T) {
	p := New()
	s := p.AddSlide()
	s.SetBackground(BackgroundOptions{Gradient: &GradientFill{
		Stops: []GradientStop{{Pos: 0, Color: "F6F2E9"}, {Pos: 100, Color: "EFE8D8"}},
		Angle: 135,
	}})
	xml := s.generateSlide()
	if !strings.Contains(xml, `<p:bg><p:bgPr><a:gradFill`) {
		t.Error("背景渐变没有写入")
	}
}
