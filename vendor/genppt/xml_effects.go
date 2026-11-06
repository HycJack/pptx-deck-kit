package genppt

import (
	"strings"
)

// 本文件集中处理「视觉特效」与「动画」两类 OOXML：
//   a:gradFill / a:ln / a:effectLst  —— 形状的填充、边框、阴影
//   p:transition                     —— 幻灯片切换
//   p:timing                         —— 入场动画时间轴
//
// 设计原则：所有写入都走显式节点。缺节点时 PowerPoint 会回落到主题默认值
// （例如描边变成主题色 2F5496），这是 genppt 早期「细线莫名变蓝」的根因。

// writeFill 写入填充：渐变优先于纯色。
func writeFill(sb *strings.Builder, solid string, transparency float64, g *GradientFill) {
	if g != nil && len(g.Stops) >= 2 {
		sb.WriteString(`<a:gradFill rotWithShape="1"><a:gsLst>`)
		for _, st := range g.Stops {
			sb.WriteString(`<a:gs pos="`)
			sb.WriteString(itoa(int(st.Pos * 1000)))
			sb.WriteString(`"><a:srgbClr val="`)
			sb.WriteString(ParseColor(st.Color))
			sb.WriteString(`"`)
			if st.Alpha > 0 && st.Alpha < 100 {
				sb.WriteString(`><a:alpha val="`)
				sb.WriteString(itoa(int(st.Alpha * 1000)))
				sb.WriteString(`"/></a:srgbClr>`)
			} else {
				sb.WriteString(`/>`)
			}
			sb.WriteString(`</a:gs>`)
		}
		sb.WriteString(`</a:gsLst>`)
		if g.Radial {
			sb.WriteString(`<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>`)
		} else {
			sb.WriteString(`<a:lin ang="`)
			sb.WriteString(itoa(int(g.Angle * 60000)))
			sb.WriteString(`" scaled="0"/>`)
		}
		sb.WriteString(`</a:gradFill>`)
		return
	}

	if solid != "" {
		sb.WriteString(`<a:solidFill><a:srgbClr val="`)
		sb.WriteString(ParseColor(solid))
		sb.WriteString(`"`)
		if transparency > 0 {
			sb.WriteString(`><a:alpha val="`)
			sb.WriteString(itoa(int((100 - transparency) * 1000)))
			sb.WriteString(`"/></a:srgbClr>`)
		} else {
			sb.WriteString(`/>`)
		}
		sb.WriteString(`</a:solidFill>`)
		return
	}

	sb.WriteString(`<a:noFill/>`)
}

// writeLine 写入边框。无边框时必须显式写 <a:noFill/>，
// 否则 PowerPoint 会套用主题默认描边（2F5496 蓝）。
func writeLine(sb *strings.Builder, o ShapeOptions) {
	noBorder := o.LineStyle == BorderNone || o.LineColor == "none"
	if !noBorder && o.LineColor != "" && o.LineWidth > 0 {
		sb.WriteString(`<a:ln w="`)
		sb.WriteString(itoa(int(o.LineWidth * 12700)))
		sb.WriteString(`"><a:solidFill><a:srgbClr val="`)
		sb.WriteString(ParseColor(o.LineColor))
		sb.WriteString(`"/></a:solidFill>`)
		switch o.LineStyle {
		case BorderDash:
			sb.WriteString(`<a:prstDash val="dash"/>`)
		case BorderDot:
			sb.WriteString(`<a:prstDash val="sysDot"/>`)
		}
		sb.WriteString(`</a:ln>`)
		return
	}
	sb.WriteString(`<a:ln><a:noFill/></a:ln>`)
}

// writeShadow 写入阴影。ShadowFX 优先，Shadow=true 保留旧的固定参数。
func writeShadow(sb *strings.Builder, o ShapeOptions) {
	if o.ShadowFX != nil {
		s := o.ShadowFX
		blur, dist := s.Blur, s.Dist
		if blur <= 0 {
			blur = 4
		}
		color := s.Color
		if color == "" {
			color = "000000"
		}
		alpha := s.Alpha
		if alpha <= 0 {
			alpha = 18
		}
		sb.WriteString(`<a:effectLst><a:outerShdw blurRad="`)
		sb.WriteString(itoa(int(blur * 12700)))
		sb.WriteString(`" dist="`)
		sb.WriteString(itoa(int(dist * 12700)))
		sb.WriteString(`" dir="`)
		sb.WriteString(itoa(int(s.Direction * 60000)))
		sb.WriteString(`" algn="ctr" rotWithShape="0"><a:srgbClr val="`)
		sb.WriteString(ParseColor(color))
		sb.WriteString(`"><a:alpha val="`)
		sb.WriteString(itoa(int(alpha * 1000)))
		sb.WriteString(`"/></a:srgbClr></a:outerShdw></a:effectLst>`)
		return
	}
	if o.Shadow {
		sb.WriteString(`<a:effectLst><a:outerShdw blurRad="50800" dist="38100" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="40000"/></a:srgbClr></a:outerShdw></a:effectLst>`)
	}
}

// generateTransition 生成幻灯片切换。
// 位置有讲究：必须是 cSld、clrMapOvr 之后的第一个元素。
func (s *Slide) generateTransition() string {
	tag, ok := transitionTags[s.transition]
	if !ok {
		return ""
	}
	return `<p:transition spd="med">` + tag + `</p:transition>`
}

// mediaPlay 需要自动播放的媒体
type mediaPlay struct {
	id      int
	isVideo bool
	loop    bool
}

// collectAutoPlayMedia 找出 AutoPlay 的媒体对象（spid = 下标+2）
func (s *Slide) collectAutoPlayMedia() []mediaPlay {
	var out []mediaPlay
	for i, obj := range s.objects {
		switch o := obj.(type) {
		case *videoObject:
			if o.options.AutoPlay {
				out = append(out, mediaPlay{i + 2, true, o.options.Loop})
			}
		case *audioObject:
			if o.options.AutoPlay {
				out = append(out, mediaPlay{i + 2, false, o.options.Loop})
			}
		}
	}
	return out
}

// writeMediaPar 写入一个媒体播放行为
func writeMediaPar(sb *strings.Builder, tn *int, m mediaPlay) {
	sb.WriteString(`<p:par><p:cTn id="`)
	sb.WriteString(itoa(*tn))
	*tn++
	sb.WriteString(`" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`)
	sb.WriteString(`<p:par><p:cTn id="`)
	sb.WriteString(itoa(*tn))
	*tn++
	sb.WriteString(`" presetID="1" presetClass="mediacall" presetSubtype="0" fill="hold" nodeType="afterEffect"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`)
	sb.WriteString(`<p:cmd type="call" cmd="playFrom(0.0)"><p:cBhvr><p:cTn id="`)
	sb.WriteString(itoa(*tn))
	*tn++
	sb.WriteString(`" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="`)
	sb.WriteString(itoa(m.id))
	sb.WriteString(`"/></p:tgtEl></p:cBhvr></p:cmd>`)
	sb.WriteString(`</p:childTnLst></p:cTn></p:par>`)
	sb.WriteString(`</p:childTnLst></p:cTn></p:par>`)
}

// writeAnimEffect 写入一个入场效果。
// 层级必须是 seq → click组 → 组内容器 → 效果，少一层 PowerPoint 会把
// 一个效果拆成两个（实测 8 个形状被数成 16 个动画）。
func writeAnimEffect(sb *strings.Builder, a animationEntry, tn *int, first bool) {
	spid := itoa(a.objectIndex + 2)
	presetID, subtype := a.effect.preset()
	nodeType := "withEffect"
	if first {
		nodeType = "clickEffect"
	}
	dur := a.duration
	if dur <= 0 {
		dur = 500
	}

	outer := *tn
	*tn++
	setID := *tn
	*tn++

	sb.WriteString(`<p:par><p:cTn id="`)
	sb.WriteString(itoa(outer))
	sb.WriteString(`" presetID="`)
	sb.WriteString(itoa(presetID))
	sb.WriteString(`" presetClass="entr" presetSubtype="`)
	sb.WriteString(itoa(subtype))
	sb.WriteString(`" fill="hold" grpId="0" nodeType="`)
	sb.WriteString(nodeType)
	sb.WriteString(`"><p:stCondLst><p:cond delay="`)
	sb.WriteString(itoa(a.delay))
	sb.WriteString(`"/></p:stCondLst><p:childTnLst>`)

	// 先把可见性设成 visible
	sb.WriteString(`<p:set><p:cBhvr><p:cTn id="`)
	sb.WriteString(itoa(setID))
	sb.WriteString(`" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`)
	sb.WriteString(`<p:tgtEl><p:spTgt spid="`)
	sb.WriteString(spid)
	sb.WriteString(`"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>`)
	sb.WriteString(`</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`)

	switch a.effect {
	case AnimAppear:
		// 只需要可见性，没有过渡行为
	case AnimWipe:
		sb.WriteString(`<p:animEffect transition="in" filter="wipe(up)"><p:cBhvr><p:cTn id="`)
		sb.WriteString(itoa(*tn))
		*tn++
		sb.WriteString(`" dur="`)
		sb.WriteString(itoa(dur))
		sb.WriteString(`"/><p:tgtEl><p:spTgt spid="`)
		sb.WriteString(spid)
		sb.WriteString(`"/></p:tgtEl></p:cBhvr></p:animEffect>`)
	case AnimFly:
		// 位置动画：ppt_x 从左侧 1/3 处回到原位
		sb.WriteString(`<p:anim calcmode="lin" valueType="num"><p:cBhvr additive="base"><p:cTn id="`)
		sb.WriteString(itoa(*tn))
		*tn++
		sb.WriteString(`" dur="`)
		sb.WriteString(itoa(dur))
		sb.WriteString(`" fill="hold"/><p:tgtEl><p:spTgt spid="`)
		sb.WriteString(spid)
		sb.WriteString(`"/></p:tgtEl><p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst></p:cBhvr>`)
		sb.WriteString(`<p:tavLst><p:tav tm="0"><p:val><p:strVal val="#ppt_x-1/3"/></p:val></p:tav><p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav></p:tavLst></p:anim>`)
		// 配合淡入，避免纯位移显得硬
		sb.WriteString(`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="`)
		sb.WriteString(itoa(*tn))
		*tn++
		sb.WriteString(`" dur="300"/><p:tgtEl><p:spTgt spid="`)
		sb.WriteString(spid)
		sb.WriteString(`"/></p:tgtEl></p:cBhvr></p:animEffect>`)
	default: // AnimFade
		sb.WriteString(`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="`)
		sb.WriteString(itoa(*tn))
		*tn++
		sb.WriteString(`" dur="`)
		sb.WriteString(itoa(dur))
		sb.WriteString(`"/><p:tgtEl><p:spTgt spid="`)
		sb.WriteString(spid)
		sb.WriteString(`"/></p:tgtEl></p:cBhvr></p:animEffect>`)
	}

	sb.WriteString(`</p:childTnLst></p:cTn></p:par>`)
}

// generateAnimTiming 生成入场动画时间轴（含自动播放媒体）。
func (s *Slide) generateAnimTiming() string {
	if len(s.animations) == 0 {
		return ""
	}

	var sb strings.Builder
	sb.WriteString(`<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>`)
	// 第三层：一次点击一组
	sb.WriteString(`<p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>`)
	// 第四层：组内容器（缺了这层，效果会被拆成两个）
	sb.WriteString(`<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`)

	tn := 5
	for i, a := range s.animations {
		writeAnimEffect(&sb, a, &tn, i == 0)
	}
	for _, m := range s.collectAutoPlayMedia() {
		writeMediaPar(&sb, &tn, m)
	}

	sb.WriteString(`</p:childTnLst></p:cTn></p:par>`)
	sb.WriteString(`</p:childTnLst></p:cTn></p:par>`)
	sb.WriteString(`</p:childTnLst></p:cTn>`)
	sb.WriteString(`<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>`)
	sb.WriteString(`<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>`)
	sb.WriteString(`</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>`)

	sb.WriteString(`<p:bldLst>`)
	for _, a := range s.animations {
		sb.WriteString(`<p:bldP spid="`)
		sb.WriteString(itoa(a.objectIndex + 2))
		sb.WriteString(`" grpId="0"/>`)
	}
	sb.WriteString(`</p:bldLst></p:timing>`)

	return sb.String()
}
