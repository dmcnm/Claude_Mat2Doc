classdef CT_Document < mat2doc.oxml.BaseOxmlElement
% CT_DOCUMENT Custom element class for the <w:document> element.
%
%   The root element of a document.xml file. Registered for <w:document>
%   (docx/oxml/__init__.py:101).
%
%   THE P2-3 BYTE-RISK (boundary-audit flag): registering w:document ->
%   CT_Document (a BaseOxmlElement subclass) means document.xml now PARSES to a
%   CT_Document with the descriptor machinery, instead of a plain XmlElement, and
%   RESERIALIZES through the BaseOxmlElement path on save. This is byte-neutral:
%   BaseOxmlElement adds NO serialization override (both XmlElement and
%   BaseOxmlElement serialize via the same +oxml/serialize_part_xml walk), so the
%   frozen M1 document.xml (1548 B) does not move one byte. Re-proven by the
%   17-part M1 sweep at P2-3 Gate 1 (mat2doc.Document().save == references\s0001,
%   17/17, esp. word/document.xml byte-identical). See audit_P2-3_document_shell.md.
%
%   DESCRIPTORS: one ZeroOrOne child descriptor (document.py:18):
%     body = ZeroOrOne("w:body")     -- successors=() (default), so insertion
%                                       always APPENDS (NO_SUCCESSORS).
%   Ported per design.md section 2 as a Constant schema table (BODY_TAG) + the
%   generated one-line delegating members calling the BaseOxmlElement child
%   engine (getChild / getOrAddChild / newChild / insertChildInSequence /
%   addChild / removeChild). ZeroOrOne generates: get.body, get_or_add_body,
%   new_body_, insert_body_, add_body_, remove_body_ (xmlchemy.py 543-573).
%
%   FORWARD DEP (out of scope, resolves generic): CT_Body IS registered by P2-3
%   (sibling file), so `.body` yields a real CT_Body. The sectPr_lst xpath
%   descends into w:p/w:pPr/w:sectPr -- CT_P (P4) and CT_SectPr (P5) are NOT
%   registered yet, so those matched nodes are generic XmlElement. sectPr_lst is
%   purely xpath-driven and works now, yielding generic elements until P4/P5
%   register the richer classes (the returned node-set is identical either way --
%   only the element CLASS of the matches changes, not their identity/order).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1 contract):
%   forwards ALL positional args to BaseOxmlElement via varargin with NO nsmap
%   re-validation -- REQUIRED because the parser instantiates this class as
%   feval(cls, name, ownDecls[, resolvedUri]) with ownDecls an Nx2 decl-pair when
%   it hits the registered <w:document> root of a real document.xml.
%
%   None SENTINEL (H3): inline `isequal(x, [])` for `x is None` (established
%   Mat2Doc oxml convention -- XmlElement.m / BaseOxmlElement.m /
%   CT_CoreProperties.m). The descriptor engine handles the None/absent cases;
%   this class adds none of its own.
%
%   Example:
%       d = mat2doc.oxml.OxmlElement("w:document");   % registry -> CT_Document
%       body = d.get_or_add_body();                   % creates <w:body/> if absent
%       secs = d.sectPr_lst;                          % 1x0 typed array (none yet)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/document.py::CT_Document
%   (registered for <w:document>, oxml/__init__.py:101)

    properties (Constant, Hidden)  % schema table (from the Python ZeroOrOne declaration)
        NO_SUCCESSORS = string.empty(1, 0)   % body has successors=() -> append
        BODY_TAG = "w:body"                  % ZeroOrOne @ document.py:18
    end

    properties (Dependent)  % generated ZeroOrOne getter + the sectPr_lst @property
        body                 % <w:body> child or [] (None) if absent
        sectPr_lst           % all directly-accessible w:sectPr (document order)
    end

    methods
        function obj = CT_Document(varargin)
            % CT_DOCUMENT Construct a loose <w:document> element.
            %   TRANSPARENT PASS-THROUGH (design.md section 2 INT-1): forward all
            %   positional args verbatim; the base XmlElement ctor is the single
            %   point accepting both nsmap currencies (struct / Nx2 decl-pair).
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ============ generated ZeroOrOne descriptor members (body) ============
        % get.body (child or []), get_or_add_body, new_body_, insert_body_,
        % add_body_, remove_body_ -- one-line delegation to the BaseOxmlElement
        % child engine. successors=() (append) -> NO_SUCCESSORS shared.
        function child = get.body(obj);            child = obj.getChild(obj.BODY_TAG); end
        function child = get_or_add_body(obj);     child = obj.getOrAddChild(obj.BODY_TAG, obj.NO_SUCCESSORS); end
        function child = new_body_(obj);           child = obj.newChild(obj.BODY_TAG); end
        function child = insert_body_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_body_(obj, varargin); child = obj.addChild(obj.BODY_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function remove_body_(obj);                obj.removeChild(obj.BODY_TAG); end

        % ===================== sectPr_lst (@property) =====================
        function lst = get.sectPr_lst(obj)
            % SECTPR_LST All w:sectPr elements directly accessible from the
            %   document element, in document order (document.py 20-32). The last
            %   is always w:body/w:sectPr; all preceding are w:p/w:pPr/w:sectPr.
            %   Does NOT include a sectPr nested under an intervening w:sdt /
            %   customXml / revision-mark layer (the xpath is a fixed two-branch
            %   union, not a descendant scan).
            %
            %   H9: xpath() returns a MATERIALIZED typed array (Python returns a
            %   list); empty typed array when none present (never [] / None).
            %   H1: xpath positions are already 1-based document order -- no shift.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/document.py::
            %   CT_Document.sectPr_lst
            lst = obj.xpath("./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr");
        end
    end
end
