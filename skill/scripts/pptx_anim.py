# -*- coding: utf-8 -*-
"""
给任意 pptx 注入切换与入场动画。

背景：python-pptx 完全没有动画 API（pptx 包内 `p:timing` 仅作为 schema 排序占位出现）。
OOXML 的动画是 <p:timing> 里的一棵行为树：
    tmRoot → seq(mainSeq) → 段落组(par) → 效果组(par) → 单个行为(set/anim/animEffect)
每个行为用 spTgt@spid 指向形状 id，所以必须先建形状、后写时间轴。

用法：
    from pptx_anim import Animator
    a = Animator(slide)
    a.entrance(shape, "fade", delay=0)
    a.entrance(shape2, "fly", delay=200)
    a.build()
    a.transition("fade")
"""
from lxml import etree
from pptx.oxml.ns import nsdecls, qn

# presetID / presetSubtype 取自 PowerPoint 自己的编号
EFFECTS = {
    "fade": (10, 0),      # 淡入
    "wipe": (22, 4),      # 擦除（自下往上）
    "fly": (2, 4),        # 飞入（自左）
    "appear": (1, 0),     # 出现（无过渡）
}

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"


class Animator:
    def __init__(self, slide):
        self.slide = slide
        self._seq = []      # [(spid, effect, delay, duration)]
        self._next_id = 3

    def _id(self):
        i = self._next_id
        self._next_id += 1
        return i

    # ---------- 对外 ----------
    def entrance(self, shape, effect="fade", delay=0, duration=500):
        eff = effect.lower()
        if eff not in EFFECTS:
            raise ValueError(f"未知效果 {effect}，可选 {list(EFFECTS)}")
        self._seq.append((shape.shape_id, eff, int(delay), int(duration)))
        return self

    def transition(self, kind="fade", speed="med"):
        """kind: fade / wipe / push / cut / circle / diamond / wedge / plus / wheel / newsflash

        不含 morph：实测 PowerPoint 16.0 打开后会把 p14:morph 节点丢掉，
        与其写一个被静默忽略的节点，不如不支持。
        """
        sld = self.slide._element
        for old in sld.findall(qn("p:transition")):
            sld.remove(old)
        tags = {
            "fade": "<p:fade/>",
            "wipe": '<p:wipe dir="u"/>',
            "push": '<p:push dir="u"/>',
            "cut": "<p:cut/>",
            "circle": "<p:circle/>",
            "diamond": "<p:diamond/>",
            "wedge": "<p:wedge/>",
            "plus": "<p:plus/>",
            "wheel": '<p:wheel spokes="1"/>',
            "newsflash": "<p:newsflash/>",
        }
        inner = tags.get(kind, tags["fade"])
        xml = f'<p:transition {nsdecls("p")} spd="{speed}">{inner}</p:transition>'
        frag = etree.fromstring(xml)
        clr = sld.find(qn("p:clrMapOvr"))
        if clr is not None:
            clr.addnext(frag)
        else:
            sld.insert(0, frag)
        return self

    def build(self):
        if not self._seq:
            return self
        sld = self.slide._element
        for old in sld.findall(qn("p:timing")):
            sld.remove(old)

        root_id = 1
        seq_id = 2
        group_id = self._id()    # 外层 click 组
        inner_id = self._id()    # 组内时间容器（PowerPoint 实测需要这一层）

        effects = []
        for i, (spid, eff, delay, dur) in enumerate(self._seq):
            preset_id, subtype = EFFECTS[eff]
            node_type = "withEffect" if i else "clickEffect"
            effects.append(self._effect_xml(spid, eff, preset_id, subtype, dur,
                                            delay, node_type))

        bld = "".join(
            f'<p:bldP {nsdecls("p")} spid="{spid}" grpId="0"/>'
            for spid, *_ in self._seq
        )

        # 层级必须是 seq → click组 → 组内容器 → 效果，少一层 PowerPoint 会把
        # 一个效果拆成两个（实测 8 个形状被数成 16 个动画）。
        xml = f"""<p:timing {nsdecls("p")}>
  <p:tnLst>
    <p:par><p:cTn id="{root_id}" dur="indefinite" restart="never" nodeType="tmRoot">
      <p:childTnLst>
        <p:seq concurrent="1" nextAc="seek">
          <p:cTn id="{seq_id}" dur="indefinite" nodeType="mainSeq">
            <p:childTnLst>
              <p:par>
                <p:cTn id="{group_id}" fill="hold">
                  <p:stCondLst>
                    <p:cond delay="indefinite"/>
                    <p:cond evt="onBegin" delay="0"><p:tn val="{seq_id}"/></p:cond>
                  </p:stCondLst>
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="{inner_id}" fill="hold">
                        <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                        <p:childTnLst>
                          {''.join(effects)}
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
          <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
          <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
        </p:seq>
      </p:childTnLst>
    </p:cTn></p:par>
  </p:tnLst>
  <p:bldLst>{bld}</p:bldLst>
</p:timing>"""
        frag = etree.fromstring(xml)
        sld.append(frag)   # timing 是 CT_Slide 的最后一个元素
        return self

    # ---------- 内部 ----------
    def _effect_xml(self, spid, eff, preset_id, subtype, dur, delay, node_type):
        """一个效果 = 一个 par，里面挂 set(可见性) + 具体行为。"""
        outer = self._id()
        set_id = self._id()
        beh_id = self._id()

        set_vis = f"""<p:set>
  <p:cBhvr>
    <p:cTn id="{set_id}" dur="1" fill="hold">
      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
    </p:cTn>
    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
    <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:to><p:strVal val="visible"/></p:to>
</p:set>"""

        if eff == "appear":
            body = set_vis
        elif eff == "fade":
            body = set_vis + f"""<p:animEffect transition="in" filter="fade">
  <p:cBhvr>
    <p:cTn id="{beh_id}" dur="{dur}"/>
    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
  </p:cBhvr>
</p:animEffect>"""
        elif eff == "wipe":
            body = set_vis + f"""<p:animEffect transition="in" filter="wipe(up)">
  <p:cBhvr>
    <p:cTn id="{beh_id}" dur="{dur}"/>
    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
  </p:cBhvr>
</p:animEffect>"""
        else:  # fly：位置动画 + 淡入
            body = set_vis + f"""<p:anim calcmode="lin" valueType="num">
  <p:cBhvr additive="base">
    <p:cTn id="{beh_id}" dur="{dur}" fill="hold"/>
    <p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>
    <p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>
  </p:cBhvr>
  <p:tavLst>
    <p:tav tm="0"><p:val><p:strVal val="#ppt_x-1/3"/></p:val></p:tav>
    <p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>
  </p:tavLst>
</p:anim>
<p:animEffect transition="in" filter="fade">
  <p:cBhvr><p:cTn id="{self._id()}" dur="{min(dur, 300)}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr>
</p:animEffect>"""

        return f"""<p:par>
  <p:cTn id="{outer}" presetID="{preset_id}" presetClass="entr" presetSubtype="{subtype}"
         fill="hold" grpId="0" nodeType="{node_type}">
    <p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>
    <p:childTnLst>{body}</p:childTnLst>
  </p:cTn>
</p:par>"""
