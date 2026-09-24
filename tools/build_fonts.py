# Урезает Liberation Serif до латиницы, кириллицы и знаков препинания и кладёт в vendor/fonts.js
import base64, io
from fontTools import subset
from fontTools.ttLib import TTFont
SRC = {'serif': '/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf',
       'serifBold': '/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf'}
ranges = list(range(0x20, 0x7F)) + list(range(0xA0, 0x100)) + list(range(0x400, 0x460)) + \
         list(range(0x2010, 0x2028)) + list(range(0x2030, 0x203B)) + [0x2116, 0x20BD, 0x2212]
out = ['/* Шрифт Liberation Serif (SIL Open Font License 1.1), урезан до латиницы и кириллицы. */', 'window.SB_FONTS = {']
for key, path in SRC.items():
    f = TTFont(path)
    opts = subset.Options(); opts.hinting = False; opts.name_IDs = ['*']; opts.notdef_outline = True
    opts.layout_features = []; opts.glyph_names = False; opts.drop_tables += ['GSUB', 'GPOS', 'GDEF', 'kern', 'DSIG']
    s = subset.Subsetter(opts); s.populate(unicodes=ranges); s.subset(f)
    buf = io.BytesIO(); f.save(buf); b = buf.getvalue()
    print(key, len(b))
    out.append('  %s: "%s",' % (key, base64.b64encode(b).decode()))
out.append('};')
open('vendor/fonts.js', 'w').write('\n'.join(out) + '\n')
