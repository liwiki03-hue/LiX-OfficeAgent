import ctypes
import math
import unicodedata


def _same_line(line,glyph):
    origin,_,reference,bottom,left,right=line
    if abs(origin-glyph[1])<=glyph[3]:return True
    if glyph[2].isspace():return abs(bottom-glyph[1])<=glyph[3]
    if reference is None:return False
    limit=min(4,.35*min(reference[8],glyph[8]))
    if origin-glyph[1]>limit:return False
    if bottom-glyph[1]<=glyph[3]:return True
    overlap=min(reference[7],glyph[7])-max(reference[6],glyph[6])
    height=min(reference[7]-reference[6],glyph[7]-glyph[6])
    separate=glyph[5]<=left or glyph[4]>=right
    return height>0 and overlap>=.6*height and separate


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
            glyphs.append((x.value,y.value,' ' if value.isspace() else value,min(1.5,max(.5,size*.12)),box[0],box[2],box[1],box[3],size))
        lines=[]
        for glyph in sorted(glyphs,key=lambda item:-item[1]):
            line=lines[-1] if lines and _same_line(lines[-1],glyph) else None
            if line is None:
                line=[glyph[1],[],None,glyph[1],math.inf,-math.inf];lines.append(line)
            line[1].append(glyph)
            if not glyph[2].isspace():
                if line[2] is None or glyph[7]-glyph[6]>line[2][7]-line[2][6]:line[2]=glyph
                line[3]=min(line[3],glyph[1]);line[4]=min(line[4],glyph[4]);line[5]=max(line[5],glyph[5])
        result=[]
        for _,line,*_ in lines:
            ordered=sorted(line,key=lambda item:item[0])
            previous=None;prior_box=None;prior_glyph=None;pieces=[]
            for glyph in ordered:
                x,y,value,_,left,right,_,_,size=glyph
                if not value.isspace():
                    if previous is not None and x-previous<.05:return None
                    if prior_box and left<prior_box[1]-min(.5,min(right-left,prior_box[1]-prior_box[0])*.2):return None
                    if (prior_glyph is not None and abs(y-prior_glyph[1])>.5 and left-prior_glyph[5]>.25*max(size,prior_glyph[8])
                            and pieces and not pieces[-1].isspace()):pieces.append(' ')
                    previous=x;prior_box=(left,right)
                    prior_glyph=glyph
                pieces.append(value)
            result.append(''.join(pieces))
        return '\n'.join(result)
    except (ValueError,TypeError,AttributeError,ctypes.ArgumentError):
        return None
