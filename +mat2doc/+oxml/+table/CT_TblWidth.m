classdef CT_TblWidth < mat2doc.oxml.BaseOxmlElement
% CT_TBLWIDTH Used for `w:tblW`, `w:tcW` and others: a table-related width.
%
%   Two RequiredAttributes (table.py 405-406):
%     w    RequiredAttribute("w:w",    XsdInt)      -> double (int)
%     type RequiredAttribute("w:type", ST_TblWidth) -> string ("auto"/"dxa"/"nil"/"pct")
%
%   H6 (the width UNION, table.py 402-418): the `w` attr is nominally
%   ST_MeasurementOrPercent, but upstream types it as XsdInt because "only dxa
%   (twips) values are being used" -- verbatim upstream comment (table.py
%   402-404). The `width` computed property (table.py 408-418) realizes the
%   union semantics EXACTLY:
%     get: if type != "dxa" -> None (a pct/auto/nil width has NO EMU length);
%          else -> Twips(w) (twips read straight from @w:w as an integer count).
%     set: type := "dxa"; w := Emu(value).twips  (a Length is stored as its
%          twips count, and the type is forced to dxa).
%   So `width` is Length-or-None ONLY for the dxa case; every other @w:type
%   yields None on read -- NOT a divergence, the faithful union.
%
%   H3 (RequiredAttribute): both `w` and `type` are REQUIRED -- reading either
%   when its attribute is absent raises mat2doc:InvalidXmlError (getAttrRequired),
%   never a default; the setters always write (setAttrRequired), never remove.
%   Consequently the `width` GETTER reads `type` first, so on a malformed element
%   missing @w:type it raises InvalidXmlError (faithful: Python `self.type`
%   raises the same).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:tcW> node inside a real table cell's <w:tcPr>
%   (the only registered tag -- w:tblW stays a plain element; see registry C4),
%   so all positional args forward verbatim.
%
%   Example:
%       tcW = mat2doc.oxml.OxmlElement("w:tcW");
%       tcW.type = "dxa"; tcW.w = 2880;   % <w:tcW w:type="dxa" w:w="2880"/>
%       tcW.width                         % 1828800 EMU (a Length)
%       tcW.width = mat2doc.shared.Twips(1440);  % type->dxa, w->1440
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TblWidth
%   (lines 399-418; registered for w:tcW)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        W_ATTR    = "w:w"          % RequiredAttribute @ table.py:405
        W_TYPE    = "XsdInt"       % simple type (+oxml\+simpletypes)
        TYPE_ATTR = "w:type"       % RequiredAttribute @ table.py:406
        TYPE_TYPE = "ST_TblWidth"  % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties + computed width union
        w       % RequiredAttribute('w:w', XsdInt) -> double (int); InvalidXmlError if absent
        type    % RequiredAttribute('w:type', ST_TblWidth) -> string; InvalidXmlError if absent
        width   % @property: EMU Length when type=="dxa", else [] (None)
    end

    methods
        function obj = CT_TblWidth(varargin)
            % CT_TBLWIDTH Construct a loose table-width element -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- w (RequiredAttribute, XsdInt) ----
        function value = get.w(obj)
            value = obj.getAttrRequired(obj.W_ATTR, obj.W_TYPE);
        end
        function set.w(obj, value)
            obj.setAttrRequired(obj.W_ATTR, obj.W_TYPE, value);
        end

        % ---- type (RequiredAttribute, ST_TblWidth) ----
        function value = get.type(obj)
            value = obj.getAttrRequired(obj.TYPE_ATTR, obj.TYPE_TYPE);
        end
        function set.type(obj, value)
            obj.setAttrRequired(obj.TYPE_ATTR, obj.TYPE_TYPE, value);
        end

        % ---- width (@property, table.py 408-418) ----
        function value = get.width(obj)
            % Python (table.py 408-413):
            %   if self.type != "dxa": return None
            %   return Twips(self.w)
            if ~strcmp(obj.type, "dxa")   % Python: if self.type != "dxa"
                value = [];               % Python: return None
                return
            end
            value = mat2doc.shared.Twips(obj.w);   % Python: Twips(self.w)
        end
        function set.width(obj, value)
            % Python (table.py 415-418):
            %   self.type = "dxa"
            %   self.w = Emu(value).twips
            obj.type = "dxa";
            obj.w = mat2doc.shared.Emu(value).twips;
        end
    end
end
