classdef CT_LsdException < mat2doc.oxml.BaseOxmlElement
% CT_LSDEXCEPTION `<w:lsdException>` element: override visibility behaviors for a named latent style.
%
%   A child of `<w:latentStyles>`. Carries a REQUIRED `w:name` attribute
%   identifying the latent style it overrides, plus five OPTIONAL override
%   attributes (locked/qFormat/semiHidden/uiPriority/unhideWhenUsed). It has NO
%   child elements.
%
%   H3 (tri-state) for the five OPTIONAL attributes: each OptionalAttribute has
%   the DEFAULT None ([]) -- absent -> [] on get; set-to-[] removes; set-to-
%   True/False writes ST_OnOff.to_xml ("1"/"0") since [] (not True) is the
%   default (see BaseOxmlElement setAttrTyped D-delta-1). uiPriority is an
%   ST_DecimalNumber (int) optional attribute. `w:name` is REQUIRED (ST_String):
%   getAttrRequired raises mat2doc:InvalidXmlError when absent.
%
%   on_off_prop / set_on_off_prop take an ATTR NAME string and dispatch to the
%   named property (Python getattr/setattr -> obj.(attr_name)); LatentStyle
%   (P4-7a) calls these with "semiHidden"/"unhideWhenUsed"/"qFormat"/"locked".
%
%   delete(): see the DELETE / handle-destructor note on the method below
%   (this is the FIRST Python `delete()` element method ported in the project).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the many <w:lsdException> nodes inside styles.xml.
%
%   Example:
%       lsd = mat2doc.oxml.OxmlElement("w:lsdException");
%       lsd.name = "Normal";          % REQUIRED @w:name
%       lsd.semiHidden = true;        % <w:lsdException w:name="Normal" w:semiHidden="1"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::CT_LsdException
%   (lines 67-89; registered for w:lsdException)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        LOCKED_ATTR         = "w:locked"          % OptionalAttribute ST_OnOff        @ styles.py:71
        LOCKED_TYPE         = "ST_OnOff"
        NAME_ATTR           = "w:name"            % RequiredAttribute ST_String       @ styles.py:72
        NAME_TYPE           = "ST_String"
        QFORMAT_ATTR        = "w:qFormat"         % OptionalAttribute ST_OnOff        @ styles.py:73
        QFORMAT_TYPE        = "ST_OnOff"
        SEMIHIDDEN_ATTR     = "w:semiHidden"      % OptionalAttribute ST_OnOff        @ styles.py:74
        SEMIHIDDEN_TYPE     = "ST_OnOff"
        UIPRIORITY_ATTR     = "w:uiPriority"      % OptionalAttribute ST_DecimalNumber@ styles.py:75
        UIPRIORITY_TYPE     = "ST_DecimalNumber"
        UNHIDEWHENUSED_ATTR = "w:unhideWhenUsed"  % OptionalAttribute ST_OnOff        @ styles.py:76
        UNHIDEWHENUSED_TYPE = "ST_OnOff"
    end

    properties (Dependent)  % generated attribute descriptors
        locked          % OptionalAttribute('w:locked', ST_OnOff) -> logical or []
        name            % RequiredAttribute('w:name', ST_String) -> string; InvalidXmlError if absent
        qFormat         % OptionalAttribute('w:qFormat', ST_OnOff) -> logical or []
        semiHidden      % OptionalAttribute('w:semiHidden', ST_OnOff) -> logical or []
        uiPriority      % OptionalAttribute('w:uiPriority', ST_DecimalNumber) -> double or []
        unhideWhenUsed  % OptionalAttribute('w:unhideWhenUsed', ST_OnOff) -> logical or []
    end

    methods
        function obj = CT_LsdException(varargin)
            % CT_LSDEXCEPTION Construct a loose <w:lsdException> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- attribute get/set (styles.py 71-76) ----
        function v = get.locked(obj);            v = obj.getAttrTyped(obj.LOCKED_ATTR, obj.LOCKED_TYPE); end
        function set.locked(obj, v);             obj.setAttrTyped(obj.LOCKED_ATTR, obj.LOCKED_TYPE, v); end

        function v = get.name(obj);              v = obj.getAttrRequired(obj.NAME_ATTR, obj.NAME_TYPE); end
        function set.name(obj, v);               obj.setAttrRequired(obj.NAME_ATTR, obj.NAME_TYPE, v); end

        function v = get.qFormat(obj);           v = obj.getAttrTyped(obj.QFORMAT_ATTR, obj.QFORMAT_TYPE); end
        function set.qFormat(obj, v);            obj.setAttrTyped(obj.QFORMAT_ATTR, obj.QFORMAT_TYPE, v); end

        function v = get.semiHidden(obj);        v = obj.getAttrTyped(obj.SEMIHIDDEN_ATTR, obj.SEMIHIDDEN_TYPE); end
        function set.semiHidden(obj, v);         obj.setAttrTyped(obj.SEMIHIDDEN_ATTR, obj.SEMIHIDDEN_TYPE, v); end

        function v = get.uiPriority(obj);        v = obj.getAttrTyped(obj.UIPRIORITY_ATTR, obj.UIPRIORITY_TYPE); end
        function set.uiPriority(obj, v);         obj.setAttrTyped(obj.UIPRIORITY_ATTR, obj.UIPRIORITY_TYPE, v); end

        function v = get.unhideWhenUsed(obj);    v = obj.getAttrTyped(obj.UNHIDEWHENUSED_ATTR, obj.UNHIDEWHENUSED_TYPE); end
        function set.unhideWhenUsed(obj, v);     obj.setAttrTyped(obj.UNHIDEWHENUSED_ATTR, obj.UNHIDEWHENUSED_TYPE, v); end

        function delete(obj)
            % DELETE Remove this `w:lsdException` element from the XML document.
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::
            %   CT_LsdException.delete (lines 78-80): self.getparent().remove(self).
            %
            %   VERIFY / candidate-H17 (handle-destructor collision -- FIRST
            %   Python `delete()` element method in the project). MATLAB `delete`
            %   on a handle subclass IS the destructor: MATLAB invalidates the
            %   handle after this body runs, and also calls it during garbage
            %   collection. The single python-docx caller (styles/latent.py:127
            %   LatentStyle.delete) does `self._element.delete(); self._element =
            %   None` -- it discards the element IMMEDIATELY, so the Python
            %   "object still usable after delete()" property is never relied
            %   upon, making the destructor-invalidation behavior-neutral on the
            %   USED surface. The parent-guard + try/catch make GC teardown
            %   harmless (destruction order is undefined; a loose/already-detached
            %   element must not error during teardown). DIVERGENCE (unreachable
            %   in python-docx usage): an EXPLICIT delete() on an UNPARENTED
            %   element is a no-op here vs AttributeError (None.remove) in Python.
            %   Flagged for the Auditor to confirm the naming/semantics ruling.
            p = obj.getparent();
            if ~isequal(p, []) && isvalid(p)   % Python: self.getparent().remove(self)
                try
                    p.remove(obj);
                catch
                    % swallow errors raised during MATLAB object teardown only
                end
            end
        end

        function value = on_off_prop(obj, attr_name)
            % ON_OFF_PROP Boolean value of the attribute `attr_name`, or [] (None) if not present.
            %   Ported from python-docx v1.2.0: styles.py CT_LsdException.on_off_prop
            %   (lines 82-85): return getattr(self, attr_name). Dynamic property
            %   access obj.(attr_name) is the getattr analogue (H3: [] when absent).
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_LsdException
                attr_name (1,1) string
            end
            value = obj.(attr_name);
        end

        function set_on_off_prop(obj, attr_name, value)
            % SET_ON_OFF_PROP Set the on/off attribute `attr_name` to `value`.
            %   Ported from python-docx v1.2.0: styles.py CT_LsdException.set_on_off_prop
            %   (lines 87-89): setattr(self, attr_name, value).
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_LsdException
                attr_name (1,1) string
                value
            end
            obj.(attr_name) = value;
        end
    end
end
