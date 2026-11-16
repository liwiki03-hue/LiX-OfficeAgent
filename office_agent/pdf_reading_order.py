import ctypes
import math
import unicodedata


def reading_text(textpage,text,count,width,height):
    if len(text)!=count:return None
    from pypdfium2 import raw
    glyphs=[]
    try:
        for index,value in enumerate(text):
            if value in '\r\n\t':continue
            if unicodedata.bidirectional(value) in {'R','AL','AN','RLE','RLO','RLI','LRE','LRO','LRI','FSI','PDI','PDF','BN'}:return None
            if not value.isprintable() and not value.isspace():return None
            x,y=ctypes.c_double(),ctypes.c_double()
            if not raw.FPDFText_GetCharOrigin(textpage.raw,index,ctypes.byref(x),ctypes.byref(y)):
                if value.isspace():continue
                return None
            angle=raw.FPDFText_GetCharAngle(textpage.raw,index)
            size=raw.FPDFText_GetFontSize(textpage.raw,index)
            if (not all(math.isfinite(number) for number in (x.value,y.value,angle,size))
                    or abs(math.sin(angle))>.001 or math.cos(angle)<.999999 or not -1<=x.value<=width+1 or not -1<=y.value<=height+1):return None
            box=textpage.get_charbox(index) if not value.isspace() else (x.value,y.value,x.value,y.value)
            glyphs.append((x.value,y.value,' ' if value.isspace() else value,min(1.5,max(.5,size*.12)),box[0],box[2]))
        lines=[]
        for glyph in sorted(glyphs,key=lambda item:-item[1]):
            line=lines[-1] if lines and abs(lines[-1][0]-glyph[1])<=glyph[3] else None
            if line is None:
                line=[glyph[1],[]];lines.append(line)
            line[1].append(glyph)
        result=[]
        for _,line in lines:
            ordered=sorted(line,key=lambda item:item[0])
            previous=None;prior_box=None
            for x,_,value,_,left,right in ordered:
                if not value.isspace():
                    if previous is not None and x-previous<.05:return None
                    if prior_box and left<prior_box[1]-min(.5,min(right-left,prior_box[1]-prior_box[0])*.2):return None
                    previous=x;prior_box=(left,right)
            result.append(''.join(item[2] for item in ordered))
        return '\n'.join(result)
    except (ValueError,TypeError,AttributeError,ctypes.ArgumentError):
        return None
