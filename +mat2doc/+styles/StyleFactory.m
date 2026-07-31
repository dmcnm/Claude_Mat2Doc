function style = StyleFactory(style_elm)
% STYLEFACTORY Return the `Style` object of the right BaseStyle subclass.
%
%   style = MAT2DOC.STYLES.STYLEFACTORY(style_elm) returns a Style proxy of the
%   appropriate |BaseStyle| subclass for the `w:style` element `style_elm`,
%   dispatched on its `type` attribute (a WD_STYLE_TYPE member).
%
%   H10 (factory dispatch): mirrors the Python dict literal indexed by
%   `style_elm.type` (style.py 15-24):
%       {PARAGRAPH: ParagraphStyle, CHARACTER: CharacterStyle,
%        TABLE: _TableStyle,        LIST: _NumberingStyle}[style_elm.type]
%   ported as an explicit isequal/elseif chain over the four WD_STYLE_TYPE
%   members. FLAG-3 private-class mapping: _TableStyle -> TableStyle_,
%   _NumberingStyle -> NumberingStyle_.
%
%   H3 EDGE (dict KeyError on None): `style_elm.type` reads the `w:style/@w:type`
%   OptionalAttribute, which is [] (None) when the attribute is absent. Python
%   then evaluates `{...}[None]`, raising `KeyError: None` (None is not a dict
%   key). This is faithfully reproduced: a `style_elm` whose type is [] (or any
%   value not among the four members) raises mat2doc:KeyError with the Python
%   repr of the key ("None" for []). Real styles.xml `w:style` elements always
%   carry `w:type`, so this path is unreachable in normal document usage; the
%   BaseStyle.type @property (which defaults None->PARAGRAPH) is a SEPARATE read
%   and is NOT used here (style.py uses the raw oxml attribute).
%
%   NOTE (module fn): a module-level function in Python (`docx.styles.style`),
%   ported as a package function `mat2doc.styles.StyleFactory` (no class).
%
%   Example:
%       s = mat2doc.oxml.OxmlElement("w:style");
%       s.type = mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER;
%       cs = mat2doc.styles.StyleFactory(s);   % a mat2doc.styles.CharacterStyle
%
%   Ported from python-docx v1.2.0: src/docx/styles/style.py::StyleFactory
    t = style_elm.type;   % WD_STYLE_TYPE member or [] (None)
    if isequal(t, mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH)
        style = mat2doc.styles.ParagraphStyle(style_elm);
    elseif isequal(t, mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER)
        style = mat2doc.styles.CharacterStyle(style_elm);
    elseif isequal(t, mat2doc.enum.style.WD_STYLE_TYPE.TABLE)
        style = mat2doc.styles.TableStyle_(style_elm);   % Python _TableStyle
    elseif isequal(t, mat2doc.enum.style.WD_STYLE_TYPE.LIST)
        style = mat2doc.styles.NumberingStyle_(style_elm);   % Python _NumberingStyle
    else
        % Python: {...}[style_elm.type] with an unmapped key (None) -> KeyError.
        % The KeyError message is the repr of the missing key; for None -> "None".
        if isequal(t, [])
            key_repr = "None";
        else
            key_repr = string(t);   % enum member name (unreachable: all 4 mapped)
        end
        error("mat2doc:KeyError", "%s", key_repr);
    end
end
