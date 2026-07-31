classdef CT_LatentStyles < mat2doc.oxml.BaseOxmlElement
% CT_LATENTSTYLES `<w:latentStyles>` element: behavior defaults for latent styles.
%
%   Defines behavior defaults for latent styles and contains `<w:lsdException>`
%   child elements that each override those defaults for a named latent style.
%
%   Six OPTIONAL attributes (H3, each DEFAULT None ([])): count / defUIPriority
%   are ST_DecimalNumber (int); defLockedState / defQFormat / defSemiHidden /
%   defUnhideWhenUsed are ST_OnOff (bool-ish). One ZeroOrMore child descriptor
%   `lsdException` (successors=() -> APPEND; docx ZeroOrMore also generates the
%   PUBLIC add_lsdException, D-delta-4).
%
%   bool_prop(attr_name) returns |False| (not []) when the named attribute is
%   absent -- the "effective default" the LatentStyles API reads; get_by_name
%   returns the matching `<w:lsdException>` child or [] (None).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the single <w:latentStyles> node inside styles.xml.
%
%   Example:
%       ls = mat2doc.oxml.OxmlElement("w:latentStyles");
%       ls.set_bool_prop("defQFormat", true);   % <w:latentStyles w:defQFormat="1"/>
%       lsd = ls.add_lsdException(); lsd.name = "Normal";
%
%   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::CT_LatentStyles
%   (lines 33-64; registered for w:latentStyles)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        LSDEXCEPTION_TAG      = "w:lsdException"    % ZeroOrMore @ styles.py:38
        NO_SUCCESSORS         = string.empty(1, 0)  % successors=() -> APPEND
        COUNT_ATTR            = "w:count"             % OptionalAttribute ST_DecimalNumber @ styles.py:40
        COUNT_TYPE            = "ST_DecimalNumber"
        DEFLOCKEDSTATE_ATTR   = "w:defLockedState"    % OptionalAttribute ST_OnOff         @ styles.py:41
        DEFLOCKEDSTATE_TYPE   = "ST_OnOff"
        DEFQFORMAT_ATTR       = "w:defQFormat"        % OptionalAttribute ST_OnOff         @ styles.py:42
        DEFQFORMAT_TYPE       = "ST_OnOff"
        DEFSEMIHIDDEN_ATTR    = "w:defSemiHidden"     % OptionalAttribute ST_OnOff         @ styles.py:43
        DEFSEMIHIDDEN_TYPE    = "ST_OnOff"
        DEFUIPRIORITY_ATTR    = "w:defUIPriority"     % OptionalAttribute ST_DecimalNumber @ styles.py:44
        DEFUIPRIORITY_TYPE    = "ST_DecimalNumber"
        DEFUNHIDEWHENUSED_ATTR = "w:defUnhideWhenUsed" % OptionalAttribute ST_OnOff        @ styles.py:45
        DEFUNHIDEWHENUSED_TYPE = "ST_OnOff"
    end

    properties (Dependent)  % generated attribute descriptors + ZeroOrMore list getter
        count             % OptionalAttribute('w:count', ST_DecimalNumber) -> double or []
        defLockedState    % OptionalAttribute('w:defLockedState', ST_OnOff) -> logical or []
        defQFormat        % OptionalAttribute('w:defQFormat', ST_OnOff) -> logical or []
        defSemiHidden     % OptionalAttribute('w:defSemiHidden', ST_OnOff) -> logical or []
        defUIPriority     % OptionalAttribute('w:defUIPriority', ST_DecimalNumber) -> double or []
        defUnhideWhenUsed % OptionalAttribute('w:defUnhideWhenUsed', ST_OnOff) -> logical or []
        lsdException_lst  % list of <w:lsdException> children (document order)
    end

    methods
        function obj = CT_LatentStyles(varargin)
            % CT_LATENTSTYLES Construct a loose <w:latentStyles> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- attribute get/set (styles.py 40-45) ----
        function v = get.count(obj);             v = obj.getAttrTyped(obj.COUNT_ATTR, obj.COUNT_TYPE); end
        function set.count(obj, v);              obj.setAttrTyped(obj.COUNT_ATTR, obj.COUNT_TYPE, v); end

        function v = get.defLockedState(obj);    v = obj.getAttrTyped(obj.DEFLOCKEDSTATE_ATTR, obj.DEFLOCKEDSTATE_TYPE); end
        function set.defLockedState(obj, v);     obj.setAttrTyped(obj.DEFLOCKEDSTATE_ATTR, obj.DEFLOCKEDSTATE_TYPE, v); end

        function v = get.defQFormat(obj);        v = obj.getAttrTyped(obj.DEFQFORMAT_ATTR, obj.DEFQFORMAT_TYPE); end
        function set.defQFormat(obj, v);         obj.setAttrTyped(obj.DEFQFORMAT_ATTR, obj.DEFQFORMAT_TYPE, v); end

        function v = get.defSemiHidden(obj);     v = obj.getAttrTyped(obj.DEFSEMIHIDDEN_ATTR, obj.DEFSEMIHIDDEN_TYPE); end
        function set.defSemiHidden(obj, v);      obj.setAttrTyped(obj.DEFSEMIHIDDEN_ATTR, obj.DEFSEMIHIDDEN_TYPE, v); end

        function v = get.defUIPriority(obj);     v = obj.getAttrTyped(obj.DEFUIPRIORITY_ATTR, obj.DEFUIPRIORITY_TYPE); end
        function set.defUIPriority(obj, v);      obj.setAttrTyped(obj.DEFUIPRIORITY_ATTR, obj.DEFUIPRIORITY_TYPE, v); end

        function v = get.defUnhideWhenUsed(obj); v = obj.getAttrTyped(obj.DEFUNHIDEWHENUSED_ATTR, obj.DEFUNHIDEWHENUSED_TYPE); end
        function set.defUnhideWhenUsed(obj, v);  obj.setAttrTyped(obj.DEFUNHIDEWHENUSED_ATTR, obj.DEFUNHIDEWHENUSED_TYPE, v); end

        % ---- lsdException (ZeroOrMore, successors=() -> APPEND) ----
        function lst = get.lsdException_lst(obj);          lst = obj.getChildList(obj.LSDEXCEPTION_TAG); end
        function child = new_lsdException_(obj);           child = obj.newChild(obj.LSDEXCEPTION_TAG); end
        function child = insert_lsdException_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_lsdException_(obj, varargin); child = obj.addChild(obj.LSDEXCEPTION_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_lsdException(obj);            child = obj.add_lsdException_(); end   % public adder (D-delta-4)

        % =================== methods (styles.py 47-64) ===================

        function value = bool_prop(obj, attr_name)
            % BOOL_PROP Boolean value of attribute `attr_name`, or |False| if not present.
            %   Ported from python-docx v1.2.0: styles.py CT_LatentStyles.bool_prop
            %   (lines 47-53): value = getattr(self, attr_name); if value is None:
            %   return False; return value. obj.(attr_name) is the getattr
            %   analogue; H3: absent -> [] (None) -> False (logical).
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_LatentStyles
                attr_name (1,1) string
            end
            value = obj.(attr_name);
            if isequal(value, [])   % Python: if value is None
                value = false;
                return
            end
        end

        function value = get_by_name(obj, name)
            % GET_BY_NAME The `w:lsdException` child having `name`, or [] (None) if not found.
            %   Ported from python-docx v1.2.0: styles.py CT_LatentStyles.get_by_name
            %   (lines 55-60): found = self.xpath('w:lsdException[@w:name="%s"]' %
            %   name); if not found: return None; return found[0]. H1: found(1) is
            %   the first match (1-based); H3: [] when no match.
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_LatentStyles
                name (1,1) string
            end
            found = obj.xpath("w:lsdException[@w:name=""" + name + """]");
            if isempty(found)   % Python: if not found
                value = [];
                return
            end
            value = found(1);   % Python: return found[0]
        end

        function set_bool_prop(obj, attr_name, value)
            % SET_BOOL_PROP Set the on/off attribute `attr_name` to `value`.
            %   Ported from python-docx v1.2.0: styles.py CT_LatentStyles.set_bool_prop
            %   (lines 62-64): setattr(self, attr_name, bool(value)).
            %
            %   Gate-2 F-1 fix (output-visible, oracle-proven): Python bool(None)
            %   is False, so `set_bool_prop(attr, None)` WRITES w:def..="0". A naive
            %   `logical(value)` maps [] (None) to an EMPTY logical, which the OnOff
            %   setAttrTyped path treats as None and REMOVES the attribute (wrong).
            %   The [] case must therefore write the logical FALSE (-> "0"), not
            %   remove. Reachable via P4-7a LatentStyles.default_to_* setters.
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_LatentStyles
                attr_name (1,1) string
                value
            end
            if isequal(value, [])   % Python: bool(None) -> False (writes "0", not remove)
                obj.(attr_name) = false;
            else
                obj.(attr_name) = logical(value);
            end
        end
    end
end
