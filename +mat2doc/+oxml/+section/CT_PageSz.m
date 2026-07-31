classdef CT_PageSz < mat2doc.oxml.BaseOxmlElement
% CT_PAGESZ `<w:pgSz>` element, defining page dimensions and orientation.
%
%   Three OptionalAttributes (section.py 86-97):
%     w      OptionalAttribute("w:w",      ST_TwipsMeasure)  -> Length, default None
%     h      OptionalAttribute("w:h",      ST_TwipsMeasure)  -> Length, default None
%     orient OptionalAttribute("w:orient", WD_ORIENTATION,
%                              default=WD_ORIENTATION.PORTRAIT) -> member
%
%   H6 (EMU/Length): w and h are Length-typed (twips). H3 (tri-state): w/h have
%   NO Python default (None -> []); getter returns [] when absent, setter removes
%   when [] (None).
%
%   H10 + H3 (enum default): orient is the ONLY attr here with a NON-None
%   default -- WD_ORIENTATION.PORTRAIT (section.py:96). So the getter returns
%   PORTRAIT when @w:orient is absent, and the setter REMOVES @w:orient when the
%   value equals PORTRAIT (the OptionalAttribute `value == self._default` delta
%   in setAttrTyped: isequal(value, default)). WD_ORIENTATION is referenced by
%   its fully qualified name so resolveTypeCls_ dispatches to +enum verbatim.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:pgSz> inside default.docx's document.xml sectPr
%   on every M1 load, so all positional args forward verbatim.
%
%   Example:
%       ps = mat2doc.oxml.OxmlElement("w:pgSz");
%       ps.orient                                    % PORTRAIT (@w:orient absent)
%       ps.w = mat2doc.shared.Twips(12240);          % <w:pgSz w:w="12240"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_PageSz
%   (lines 86-97; registered for w:pgSz)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        W_ATTR        = "w:w"                              % OptionalAttribute @ section.py:89-91
        W_TYPE        = "ST_TwipsMeasure"
        W_DEFAULT     = []                                  % Python default: None
        H_ATTR        = "w:h"                              % OptionalAttribute @ section.py:92-94
        H_TYPE        = "ST_TwipsMeasure"
        H_DEFAULT     = []                                  % Python default: None
        ORIENT_ATTR   = "w:orient"                         % OptionalAttribute @ section.py:95-97
        ORIENT_TYPE   = "mat2doc.enum.section.WD_ORIENTATION" % enum simple-type (verbatim, resolveTypeCls_)
        % NON-None default: WD_ORIENTATION.PORTRAIT (section.py:96)
        ORIENT_DEFAULT = mat2doc.enum.section.WD_ORIENTATION.PORTRAIT
    end

    properties (Dependent)  % generated descriptor properties
        w       % OptionalAttribute('w:w', ST_TwipsMeasure) -> Length or []
        h       % OptionalAttribute('w:h', ST_TwipsMeasure) -> Length or []
        orient  % OptionalAttribute('w:orient', WD_ORIENTATION, default=PORTRAIT) -> member
    end

    methods
        function obj = CT_PageSz(varargin)
            % CT_PAGESZ Construct a loose <w:pgSz> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- w (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.w(obj)
            value = obj.getAttrTyped(obj.W_ATTR, obj.W_TYPE, obj.W_DEFAULT);
        end
        function set.w(obj, value)
            obj.setAttrTyped(obj.W_ATTR, obj.W_TYPE, value, obj.W_DEFAULT);
        end

        % ---- h (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.h(obj)
            value = obj.getAttrTyped(obj.H_ATTR, obj.H_TYPE, obj.H_DEFAULT);
        end
        function set.h(obj, value)
            obj.setAttrTyped(obj.H_ATTR, obj.H_TYPE, value, obj.H_DEFAULT);
        end

        % ---- orient (OptionalAttribute, WD_ORIENTATION, default PORTRAIT) ----
        function value = get.orient(obj)
            value = obj.getAttrTyped(obj.ORIENT_ATTR, obj.ORIENT_TYPE, obj.ORIENT_DEFAULT);
        end
        function set.orient(obj, value)
            obj.setAttrTyped(obj.ORIENT_ATTR, obj.ORIENT_TYPE, value, obj.ORIENT_DEFAULT);
        end
    end
end
