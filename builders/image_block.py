"""Helpers that emit JSX for image blocks."""
from __future__ import annotations

from string import Template

IMG_MARKER_TEMPLATE = Template("""(function(){
  log("[PY][m_img] $src_log inline=$inline");
  try {
    log("[DBG] typeof addFloatingImage=" + (typeof addFloatingImage)
        + " typeof addImageAtV2=" + (typeof addImageAtV2)
        + " typeof _normPath=" + (typeof _normPath));
    log("[DBG] tf=" + (tf&&tf.isValid) + " story=" + (story&&story.isValid) + " page=" + (page&&page.isValid));

    try{ if(typeof flushOverflow==="function"){ var _rs=flushOverflow(story,page,tf);
      if(_rs&&_rs.frame&&_rs.page){ page=_rs.page; tf=_rs.frame; story=tf.parentStory; curTextFrame=tf; } } }catch(_){}

    var ip=(tf&&tf.isValid)?_safeIP(tf):story.insertionPoints[-1];
    var f=_normPath("$src");
    log("[DBG] _normPath ok=" + (!!f) + " exists=" + (f&&f.exists ? "Y":"N") + " fsName=" + (f?f.fsName:"NA"));

    if(f&&f.exists){
      var spec={src:f.fsName,w:"$w",h:"$h",align:"$align",spaceBefore:$sb,spaceAfter:$sa,caption:"$caption",
                inline:"$inline",wrap:"$wrap",posH:"$posH",posV:"$posV",offX:"$offX",offY:"$offY",
                distT:"$distT",distB:"$distB",distL:"$distL",distR:"$distR",forceBlock:true};
      var inl=_trim(spec.inline);
      log("[IMG-DISPATCH] src="+spec.src+" inline="+inl+" posH="+(spec.posH||"")+" posV="+(spec.posV||""));

      if(inl==="0"||/^false/i.test(inl)){
        log("[DBG] dispatch -> addFloatingImage");
        var rect=addFloatingImage(tf,story,page,spec);
        if(rect&&rect.isValid) log("[IMG] ok (float): " + spec.src);
      } else {
        log("[DBG] dispatch -> addImageAtV2");
        var rect=addImageAtV2(ip,spec);
        if(rect&&rect.isValid) log("[IMG] ok (inline): " + spec.src);
      }
    } else {
      log("[IMG] missing: $src_log");
    }
  } catch(e) {
    log("[IMG][EXC] " + e);
  }
})();""")


IMG_XML_TEMPLATE = Template("""(function(){
  log("[PY][m_xmli] $src_log");
  try{ if(typeof flushOverflow==="function"){ var _rs=flushOverflow(story,page,tf);
  if(_rs&&_rs.frame&&_rs.page){ page=_rs.page; tf=_rs.frame; story=tf.parentStory; curTextFrame=tf; } } }catch(_){}
  var ip=(tf&&tf.isValid)?_safeIP(tf):story.insertionPoints[-1];
  try{
    var para=ip.paragraphs[0]; var p0=(para&&para.isValid)?para.insertionPoints[0]:null;
    var h0=(p0&&p0.isValid&&p0.parentTextFrames&&p0.parentTextFrames.length)?p0.parentTextFrames[0]:null;
    if(h0&&h0.isValid&&tf&&tf.isValid&&h0.id!==tf.id){ ip.contents="\\r"; try{story.recompose();}catch(__){} ip=tf.insertionPoints[-1]; }
  }catch(__){}
  var f=_normPath("$src");
  if(f&&f.exists){
    var spec={src:f.fsName,w:"$w",h:"$h",align:"$align",spaceBefore:$sb,spaceAfter:$sa,caption:"$caption",
              inline:"$inline",wrap:"$wrap",posH:"$posH",posV:"$posV",offX:"$offX",offY:"$offY",
              distT:"$distT",distB:"$distB",distL:"$distL",distR:"$distR",forceBlock:true};
    var inl=_trim(spec.inline);
    if(inl==="0"||/^false/i.test(inl)){
      var rect=addFloatingImage(tf,story,page,spec);
      if(rect&&rect.isValid) log("[IMG] ok (float): "+spec.src);
    } else {
      var rect=addImageAtV2(ip,spec);
      if(rect&&rect.isValid) log("[IMG] ok (inline): "+spec.src);
    }
  } else {
    log("[IMG] missing: $src_log");
  }
})();""")


def append_marker_image(add_lines: list[str], **fields: str) -> None:
    add_lines.append(IMG_MARKER_TEMPLATE.substitute(**fields))


def append_xml_image(add_lines: list[str], **fields: str) -> None:
    add_lines.append(IMG_XML_TEMPLATE.substitute(**fields))
