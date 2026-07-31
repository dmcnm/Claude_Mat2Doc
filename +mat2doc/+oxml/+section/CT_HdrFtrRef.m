classdef CT_HdrFtrRef < mat2doc.oxml.BaseOxmlElement
% CT_HDRFTRREF `w:headerReference` and `w:footerReference` elements.
%
%   Two RequiredAttributes (section.py 54-57):
%     type_ RequiredAttribute("w:type", WD_HEADER_FOOTER) -> a member
%     rId   RequiredAttribute("r:id",   XsdString)        -> string
%
%   Python names the first member `type_` (a trailing underscore already, to
%   dodge the `type` builtin) -- NO underscore rotation applies (rotation is for
%   LEADING underscores only), so the MATLAB member is `type_` verbatim.
%
%   H5 (relationships): rId is ported here as a PLAIN r:id string attribute only.
%   The separate-part hdr/ftr relationship wiring (add_header_part etc.) is
%   P5-3b -- this class does NOT resolve or create relationships.
%
%   H10 (enum dispatch): WD_HEADER_FOOTER is python-docx's alias of
%   WD_HEADER_FOOTER_INDEX (section.py:27); the MATLAB alias class
%   mat2doc.enum.section.WD_HEADER_FOOTER forwards from_xml/to_xml to the
%   canonical enum, so getAttrRequired/setAttrRequired dispatch through it
%   (resolveTypeCls_ passes the fully qualified name verbatim, as CT_Highlight
%   does for WD_COLOR_INDEX).
%
%   H3 (RequiredAttribute): both attrs are REQUIRED -- a missing @w:type or
%   @r:id on load raises mat2doc:InvalidXmlError (getAttrRequired), never a
%   default; the setters never remove (setAttrRequired), and the docx delta's
%   to_xml==None ValueError guard applies verbatim.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:headerReference>/<w:footerReference> nodes inside a
%   real sectPr, so all positional args forward verbatim.
%
%   Example:
%       ref = mat2doc.oxml.OxmlElement("w:headerReference");   % a CT_HdrFtrRef
%       ref.type_ = mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY;
%       ref.rId = "rId7";   % <w:headerReference w:type="default" r:id="rId7"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/section.py::CT_HdrFtrRef
%   (lines 51-57; registered for w:headerReference and w:footerReference)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        TYPE_ATTR = "w:type"                              % RequiredAttribute @ section.py:54-56
        TYPE_TYPE = "mat2doc.enum.section.WD_HEADER_FOOTER" % enum simple-type (verbatim, resolveTypeCls_)
        RID_ATTR  = "r:id"                                % RequiredAttribute @ section.py:57
        RID_TYPE  = "XsdString"                           % simple type (+oxml\+simpletypes)
    end

    properties (Dependent)  % generated descriptor properties
        type_   % RequiredAttribute('w:type', WD_HEADER_FOOTER) -> member; InvalidXmlError if absent
        rId     % RequiredAttribute('r:id', XsdString) -> string; InvalidXmlError if absent
    end

    methods
        function obj = CT_HdrFtrRef(varargin)
            % CT_HDRFTRREF Construct a loose header/footer reference -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- type_ (RequiredAttribute, WD_HEADER_FOOTER) ----
        function value = get.type_(obj)
            value = obj.getAttrRequired(obj.TYPE_ATTR, obj.TYPE_TYPE);
        end
        function set.type_(obj, value)
            obj.setAttrRequired(obj.TYPE_ATTR, obj.TYPE_TYPE, value);
        end

        % ---- rId (RequiredAttribute, XsdString) ----
        function value = get.rId(obj)
            value = obj.getAttrRequired(obj.RID_ATTR, obj.RID_TYPE);
        end
        function set.rId(obj, value)
            obj.setAttrRequired(obj.RID_ATTR, obj.RID_TYPE, value);
        end
    end
end
