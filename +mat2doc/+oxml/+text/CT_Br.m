classdef CT_Br < mat2doc.oxml.BaseOxmlElement
% CT_BR `<w:br>` element, indicating a line, page, or column break in a run.
%
%   Registered for <w:br> (docx/oxml/__init__.py:72).
%
%   DESCRIPTORS (run.py 174-177):
%     type  = OptionalAttribute("w:type", ST_BrType, default="textWrapping")
%     clear = OptionalAttribute("w:clear", ST_BrClear)          (default None -> [])
%
%   H3 tri-state:
%     * `type` has a NON-None default "textWrapping": getAttrTyped returns
%       "textWrapping" when @w:type is absent; the setter (D-delta-1) removes
%       @w:type when set to None ([]) OR to the default "textWrapping".
%     * `clear` has no default (None -> []): absent -> []; set [] removes @w:clear.
%   ST_BrType / ST_BrClear are XsdString simple types (P3-2); referenced by bare
%   short name (resolveTypeCls_ prefixes +oxml.simpletypes).
%
%   __str__ (run.py 179-188) -> str_(): text equivalent of the break. A LINE
%   break ("textWrapping", incl. the absent-default case) maps to "\n"; column
%   and page breaks map to "". H2: MATLAB "\n" in a double-quoted literal is the
%   two chars backslash-n, so the newline is produced via newline (char 10).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Example:
%       br = mat2doc.oxml.OxmlElement("w:br");
%       br.str_()               % "\n" (LF; type defaults to "textWrapping")
%       br.type = "page";       % <w:br w:type="page"/>
%       br.str_()               % ""
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_Br
%   (lines 171-188; registered for w:br)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        TYPE_ATTR     = "w:type"          % OptionalAttribute @ run.py:174-176
        TYPE_TYPE     = "ST_BrType"       % XsdString simple type (P3-2)
        TYPE_DEFAULT  = "textWrapping"    % Python default="textWrapping" (NON-None)
        CLEAR_ATTR    = "w:clear"         % OptionalAttribute @ run.py:177
        CLEAR_TYPE    = "ST_BrClear"      % XsdString simple type (P3-2)
        CLEAR_DEFAULT = []                % Python default: None
    end

    properties (Dependent)  % generated OptionalAttribute properties
        type     % "page"|"column"|"textWrapping"; "textWrapping" when @w:type absent
        clear    % "none"|"left"|"right"|"all", or [] when @w:clear absent
    end

    methods
        function obj = CT_Br(varargin)
            % CT_BR Construct a loose <w:br> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.type(obj)
            value = obj.getAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE, obj.TYPE_DEFAULT);
        end
        function set.type(obj, value)
            obj.setAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE, value, obj.TYPE_DEFAULT);
        end

        function value = get.clear(obj)
            value = obj.getAttrTyped(obj.CLEAR_ATTR, obj.CLEAR_TYPE, obj.CLEAR_DEFAULT);
        end
        function set.clear(obj, value)
            obj.setAttrTyped(obj.CLEAR_ATTR, obj.CLEAR_TYPE, value, obj.CLEAR_DEFAULT);
        end

        function value = str_(obj)
            % STR_ Text equivalent (run.py 179-188): "\n" if line break else "".
            %   Python: return "\n" if self.type == "textWrapping" else ""
            if obj.type == "textWrapping"
                value = string(newline);   % actual LF (char 10); NOT the literal "\n" (H2)
            else
                value = "";
            end
        end
    end
end
