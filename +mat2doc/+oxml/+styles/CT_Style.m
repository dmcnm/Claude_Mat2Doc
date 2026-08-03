classdef CT_Style < mat2doc.oxml.BaseOxmlElement
% CT_STYLE A `<w:style>` element, representing a style definition.
%
%   Carries four OPTIONAL attributes (type -> WD_STYLE_TYPE enum, styleId ->
%   ST_String, default / customStyle -> ST_OnOff) and TEN ZeroOrOne child
%   descriptors. The `.._val` @property members read/write the `w:val`
%   attribute of the corresponding child element.
%
%   H11 (child ordering -- CRITICAL): _tag_seq (styles.py 95-118, VERBATIM, 22
%   tags) is stored as the Constant TAG_SEQ. Each ZeroOrOne descriptor's
%   successors = Python _tag_seq[s0:] is TAG_SEQ(s0+1:end) (H1: 0-based Python
%   slice start s0 -> 1-based MATLAB start s0+1). insert_element_before scans
%   these successors in argument order and inserts the new child before the
%   FIRST present successor tag, re-sorting scrambled adds into canonical schema
%   order so styles.xml stays byte-identical.
%
%   H11 SUCCESSOR-SLICE TABLE (descriptor : own 1-based idx : Python slice : MATLAB slice):
%     name           =  1 : successors=_tag_seq[1:]  -> TAG_SEQ(2:end)
%     basedOn        =  3 : successors=_tag_seq[3:]  -> TAG_SEQ(4:end)   (aliases@2 skipped)
%     next           =  4 : successors=_tag_seq[4:]  -> TAG_SEQ(5:end)
%     uiPriority     =  8 : successors=_tag_seq[8:]  -> TAG_SEQ(9:end)   (link/autoRedefine/hidden@5-7 skipped)
%     semiHidden     =  9 : successors=_tag_seq[9:]  -> TAG_SEQ(10:end)
%     unhideWhenUsed = 10 : successors=_tag_seq[10:] -> TAG_SEQ(11:end)
%     qFormat        = 11 : successors=_tag_seq[11:] -> TAG_SEQ(12:end)
%     locked         = 12 : successors=_tag_seq[12:] -> TAG_SEQ(13:end)
%     pPr            = 17 : successors=_tag_seq[17:] -> TAG_SEQ(18:end)  (personal*/rsid@13-16 skipped)
%     rPr            = 18 : successors=_tag_seq[18:] -> TAG_SEQ(19:end)
%
%   NOTE (no `link` descriptor): python-docx v1.2.0 lists w:link only in
%   _tag_seq (position 5); it declares NO `link` ZeroOrOne descriptor (nor
%   aliases/autoRedefine/hidden/personal*/rsid/tblPr/trPr/tcPr/tblStylePr). This
%   port faithfully has NO link accessor -- w:link parses as a generic
%   XmlElement (unregistered) and round-trips byte-identically.
%
%   GENERATED DESCRIPTOR FAMILY (per ZeroOrOne, docx xmlchemy form): get.x,
%   get_or_add_x, new_x_, insert_x_, add_x_, remove_x_ (underscore rotation:
%   Python _new_x/_insert_x/_add_x/_remove_x -> new_x_/insert_x_/add_x_/
%   remove_x_; get_or_add_x public). ALL 10 descriptors use the generic
%   BaseOxmlElement engine (no _new_x/_insert_x override on CT_Style).
%
%   CHILD-CLASS REGISTRATION (parse-time class of each descriptor's child, all
%   registered by THIS WP's styles block or earlier): name/basedOn/next ->
%   CT_String; uiPriority -> CT_DecimalNumber; semiHidden/unhideWhenUsed/qFormat/
%   locked -> CT_OnOff; pPr -> CT_PPr (P4-2); rPr -> CT_RPr (P4-1a).
%
%   H3 (tri-state): basedOn_val/name_val/uiPriority_val return [] (None) when
%   absent; locked_val/qFormat_val/semiHidden_val/unhideWhenUsed_val return the
%   logical FALSE when absent (Python returns bool False, styles.py 172-269).
%   The four attributes each have DEFAULT None ([]).
%
%   H4 (truthiness): the boolean-ish setters port `if bool(value) is True:` /
%   `if bool(value):` as `~isequal(value, []) && value`.
%
%   delete(): see the DELETE / handle-destructor note on the method below.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the 164 <w:style> nodes inside a real styles.xml.
%
%   Example:
%       s = mat2doc.oxml.OxmlElement("w:style");
%       s.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
%       s.styleId = "Heading1"; s.name_val = "heading 1";
%
%   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::CT_Style
%   (lines 92-269; registered for w:style)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (styles.py 95-118; 22 tags) + attr schema
        TAG_SEQ = [ ...
            "w:name", "w:aliases", "w:basedOn", "w:next", ...
            "w:link", "w:autoRedefine", "w:hidden", "w:uiPriority", ...
            "w:semiHidden", "w:unhideWhenUsed", "w:qFormat", "w:locked", ...
            "w:personal", "w:personalCompose", "w:personalReply", "w:rsid", ...
            "w:pPr", "w:rPr", "w:tblPr", "w:trPr", ...
            "w:tcPr", "w:tblStylePr" ]
        TYPE_ATTR        = "w:type"                          % OptionalAttribute WD_STYLE_TYPE @ styles.py:131
        TYPE_TYPE        = "mat2doc.enum.style.WD_STYLE_TYPE" % enum simple-type (verbatim, resolveTypeCls_)
        STYLEID_ATTR     = "w:styleId"                       % OptionalAttribute ST_String     @ styles.py:134
        STYLEID_TYPE     = "ST_String"
        DEFAULT_ATTR     = "w:default"                       % OptionalAttribute ST_OnOff      @ styles.py:137
        DEFAULT_TYPE     = "ST_OnOff"
        CUSTOMSTYLE_ATTR = "w:customStyle"                   % OptionalAttribute ST_OnOff      @ styles.py:138
        CUSTOMSTYLE_TYPE = "ST_OnOff"
    end

    properties (Dependent)  % 10 ZeroOrOne getters + 4 attributes + @property members
        % -- 10 ZeroOrOne child getters (read-only; use get_or_add_x/remove_x_) --
        name
        basedOn
        next
        uiPriority
        semiHidden
        unhideWhenUsed
        qFormat
        locked
        pPr
        rPr
        % -- 4 attributes (styles.py 131-138) --
        type         % OptionalAttribute('w:type', WD_STYLE_TYPE) -> enum member or []
        styleId      % OptionalAttribute('w:styleId', ST_String) -> string or []
        default      % OptionalAttribute('w:default', ST_OnOff) -> logical or []
        customStyle  % OptionalAttribute('w:customStyle', ST_OnOff) -> logical or []
        % -- @property members (styles.py 140-269) --
        basedOn_val      % w:basedOn/@w:val or []
        base_style       % sibling CT_Style this is based on, or []
        locked_val       % w:locked/@w:val or FALSE
        name_val         % <w:name>/@w:val or []
        next_style       % sibling CT_Style identified by w:next/@w:val, or []
        qFormat_val      % w:qFormat/@w:val or FALSE
        semiHidden_val   % <w:semiHidden>/@w:val or FALSE
        uiPriority_val   % <w:uiPriority>/@w:val or []
        unhideWhenUsed_val % w:unhideWhenUsed/@w:val or FALSE
    end

    methods
        function obj = CT_Style(varargin)
            % CT_STYLE Construct a loose <w:style> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- name (ZeroOrOne, successors=_tag_seq[1:] -> TAG_SEQ(2:end)) ----
        function child = get.name(obj);            child = obj.getChild("w:name"); end
        function child = get_or_add_name(obj);     child = obj.getOrAddChild("w:name", obj.TAG_SEQ(2:end)); end
        function child = new_name_(obj);           child = obj.newChild("w:name"); end
        function child = insert_name_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(2:end)); end
        function child = add_name_(obj, varargin); child = obj.addChild("w:name", obj.TAG_SEQ(2:end), varargin{:}); end
        function remove_name_(obj);                obj.removeChild("w:name"); end

        % ---- basedOn (ZeroOrOne, successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ----
        function child = get.basedOn(obj);            child = obj.getChild("w:basedOn"); end
        function child = get_or_add_basedOn(obj);     child = obj.getOrAddChild("w:basedOn", obj.TAG_SEQ(4:end)); end
        function child = new_basedOn_(obj);           child = obj.newChild("w:basedOn"); end
        function child = insert_basedOn_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_basedOn_(obj, varargin); child = obj.addChild("w:basedOn", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_basedOn_(obj);                obj.removeChild("w:basedOn"); end

        % ---- next (ZeroOrOne, successors=_tag_seq[4:] -> TAG_SEQ(5:end)) ----
        function child = get.next(obj);            child = obj.getChild("w:next"); end
        function child = get_or_add_next(obj);     child = obj.getOrAddChild("w:next", obj.TAG_SEQ(5:end)); end
        function child = new_next_(obj);           child = obj.newChild("w:next"); end
        function child = insert_next_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(5:end)); end
        function child = add_next_(obj, varargin); child = obj.addChild("w:next", obj.TAG_SEQ(5:end), varargin{:}); end
        function remove_next_(obj);                obj.removeChild("w:next"); end

        % ---- uiPriority (ZeroOrOne, successors=_tag_seq[8:] -> TAG_SEQ(9:end)) ----
        function child = get.uiPriority(obj);            child = obj.getChild("w:uiPriority"); end
        function child = get_or_add_uiPriority(obj);     child = obj.getOrAddChild("w:uiPriority", obj.TAG_SEQ(9:end)); end
        function child = new_uiPriority_(obj);           child = obj.newChild("w:uiPriority"); end
        function child = insert_uiPriority_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(9:end)); end
        function child = add_uiPriority_(obj, varargin); child = obj.addChild("w:uiPriority", obj.TAG_SEQ(9:end), varargin{:}); end
        function remove_uiPriority_(obj);                obj.removeChild("w:uiPriority"); end

        % ---- semiHidden (ZeroOrOne, successors=_tag_seq[9:] -> TAG_SEQ(10:end)) ----
        function child = get.semiHidden(obj);            child = obj.getChild("w:semiHidden"); end
        function child = get_or_add_semiHidden(obj);     child = obj.getOrAddChild("w:semiHidden", obj.TAG_SEQ(10:end)); end
        function child = new_semiHidden_(obj);           child = obj.newChild("w:semiHidden"); end
        function child = insert_semiHidden_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(10:end)); end
        function child = add_semiHidden_(obj, varargin); child = obj.addChild("w:semiHidden", obj.TAG_SEQ(10:end), varargin{:}); end
        function remove_semiHidden_(obj);                obj.removeChild("w:semiHidden"); end

        % ---- unhideWhenUsed (ZeroOrOne, successors=_tag_seq[10:] -> TAG_SEQ(11:end)) ----
        function child = get.unhideWhenUsed(obj);            child = obj.getChild("w:unhideWhenUsed"); end
        function child = get_or_add_unhideWhenUsed(obj);     child = obj.getOrAddChild("w:unhideWhenUsed", obj.TAG_SEQ(11:end)); end
        function child = new_unhideWhenUsed_(obj);           child = obj.newChild("w:unhideWhenUsed"); end
        function child = insert_unhideWhenUsed_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(11:end)); end
        function child = add_unhideWhenUsed_(obj, varargin); child = obj.addChild("w:unhideWhenUsed", obj.TAG_SEQ(11:end), varargin{:}); end
        function remove_unhideWhenUsed_(obj);                obj.removeChild("w:unhideWhenUsed"); end

        % ---- qFormat (ZeroOrOne, successors=_tag_seq[11:] -> TAG_SEQ(12:end)) ----
        function child = get.qFormat(obj);            child = obj.getChild("w:qFormat"); end
        function child = get_or_add_qFormat(obj);     child = obj.getOrAddChild("w:qFormat", obj.TAG_SEQ(12:end)); end
        function child = new_qFormat_(obj);           child = obj.newChild("w:qFormat"); end
        function child = insert_qFormat_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(12:end)); end
        function child = add_qFormat_(obj, varargin); child = obj.addChild("w:qFormat", obj.TAG_SEQ(12:end), varargin{:}); end
        function remove_qFormat_(obj);                obj.removeChild("w:qFormat"); end

        % ---- locked (ZeroOrOne, successors=_tag_seq[12:] -> TAG_SEQ(13:end)) ----
        function child = get.locked(obj);            child = obj.getChild("w:locked"); end
        function child = get_or_add_locked(obj);     child = obj.getOrAddChild("w:locked", obj.TAG_SEQ(13:end)); end
        function child = new_locked_(obj);           child = obj.newChild("w:locked"); end
        function child = insert_locked_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(13:end)); end
        function child = add_locked_(obj, varargin); child = obj.addChild("w:locked", obj.TAG_SEQ(13:end), varargin{:}); end
        function remove_locked_(obj);                obj.removeChild("w:locked"); end

        % ---- pPr (ZeroOrOne, successors=_tag_seq[17:] -> TAG_SEQ(18:end)) ----
        function child = get.pPr(obj);            child = obj.getChild("w:pPr"); end
        function child = get_or_add_pPr(obj);     child = obj.getOrAddChild("w:pPr", obj.TAG_SEQ(18:end)); end
        function child = new_pPr_(obj);           child = obj.newChild("w:pPr"); end
        function child = insert_pPr_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(18:end)); end
        function child = add_pPr_(obj, varargin); child = obj.addChild("w:pPr", obj.TAG_SEQ(18:end), varargin{:}); end
        function remove_pPr_(obj);                obj.removeChild("w:pPr"); end

        % ---- rPr (ZeroOrOne, successors=_tag_seq[18:] -> TAG_SEQ(19:end)) ----
        function child = get.rPr(obj);            child = obj.getChild("w:rPr"); end
        function child = get_or_add_rPr(obj);     child = obj.getOrAddChild("w:rPr", obj.TAG_SEQ(19:end)); end
        function child = new_rPr_(obj);           child = obj.newChild("w:rPr"); end
        function child = insert_rPr_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(19:end)); end
        function child = add_rPr_(obj, varargin); child = obj.addChild("w:rPr", obj.TAG_SEQ(19:end), varargin{:}); end
        function remove_rPr_(obj);                obj.removeChild("w:rPr"); end

        % =================== attributes (styles.py 131-138) ===================
        function v = get.type(obj);          v = obj.getAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE); end
        function set.type(obj, v);           obj.setAttrTyped(obj.TYPE_ATTR, obj.TYPE_TYPE, v); end

        function v = get.styleId(obj);       v = obj.getAttrTyped(obj.STYLEID_ATTR, obj.STYLEID_TYPE); end
        function set.styleId(obj, v);        obj.setAttrTyped(obj.STYLEID_ATTR, obj.STYLEID_TYPE, v); end

        function v = get.default(obj);       v = obj.getAttrTyped(obj.DEFAULT_ATTR, obj.DEFAULT_TYPE); end
        function set.default(obj, v);        obj.setAttrTyped(obj.DEFAULT_ATTR, obj.DEFAULT_TYPE, v); end

        function v = get.customStyle(obj);   v = obj.getAttrTyped(obj.CUSTOMSTYLE_ATTR, obj.CUSTOMSTYLE_TYPE); end
        function set.customStyle(obj, v);    obj.setAttrTyped(obj.CUSTOMSTYLE_ATTR, obj.CUSTOMSTYLE_TYPE, v); end

        % =================== @property members (styles.py 140-269) ===================

        % ---- basedOn_val (styles.py 140-153) ----
        function value = get.basedOn_val(obj)
            basedOn = obj.basedOn;
            if isequal(basedOn, [])   % Python: if basedOn is None
                value = [];
                return
            end
            value = basedOn.val;
        end
        function set.basedOn_val(obj, value)
            if isequal(value, [])   % Python: if value is None
                obj.remove_basedOn_();
            else
                b = obj.get_or_add_basedOn();
                b.val = value;
            end
        end

        % ---- base_style (styles.py 155-166) ----
        function value = get.base_style(obj)
            basedOn = obj.basedOn;
            if isequal(basedOn, [])   % Python: if basedOn is None
                value = [];
                return
            end
            styles = obj.getparent();
            base_style = styles.get_by_id(basedOn.val);
            if isequal(base_style, [])   % Python: if base_style is None
                value = [];
                return
            end
            value = base_style;
        end

        % ---- locked_val (styles.py 172-185) ----
        function value = get.locked_val(obj)
            locked = obj.locked;
            if isequal(locked, [])   % Python: if locked is None
                value = false;       % Python: return False
                return
            end
            value = locked.val;
        end
        function set.locked_val(obj, value)
            obj.remove_locked_();
            if ~isequal(value, []) && value   % Python: if bool(value) is True (H4)
                locked = obj.add_locked_();
                locked.val = value;
            end
        end

        % ---- name_val (styles.py 187-200) ----
        function value = get.name_val(obj)
            name = obj.name;
            if isequal(name, [])   % Python: if name is None
                value = [];
                return
            end
            value = name.val;
        end
        function set.name_val(obj, value)
            obj.remove_name_();
            if ~isequal(value, [])   % Python: if value is not None
                name = obj.add_name_();
                name.val = value;
            end
        end

        % ---- next_style (styles.py 202-210) ----
        function value = get.next_style(obj)
            next = obj.next;
            if isequal(next, [])   % Python: if next is None
                value = [];
                return
            end
            styles = obj.getparent();
            value = styles.get_by_id(next.val);   % Python: return styles.get_by_id(next.val) (None if not found)
        end

        % ---- qFormat_val (styles.py 212-224) ----
        function value = get.qFormat_val(obj)
            qFormat = obj.qFormat;
            if isequal(qFormat, [])   % Python: if qFormat is None
                value = false;        % Python: return False
                return
            end
            value = qFormat.val;
        end
        function set.qFormat_val(obj, value)
            obj.remove_qFormat_();
            if ~isequal(value, []) && value   % Python: if bool(value) (H4)
                obj.add_qFormat_();           % Python: self._add_qFormat() (no .val set)
            end
        end

        % ---- semiHidden_val (styles.py 226-239) ----
        function value = get.semiHidden_val(obj)
            semiHidden = obj.semiHidden;
            if isequal(semiHidden, [])   % Python: if semiHidden is None
                value = false;           % Python: return False
                return
            end
            value = semiHidden.val;
        end
        function set.semiHidden_val(obj, value)
            obj.remove_semiHidden_();
            if ~isequal(value, []) && value   % Python: if bool(value) is True (H4)
                semiHidden = obj.add_semiHidden_();
                semiHidden.val = value;
            end
        end

        % ---- uiPriority_val (styles.py 241-254) ----
        function value = get.uiPriority_val(obj)
            uiPriority = obj.uiPriority;
            if isequal(uiPriority, [])   % Python: if uiPriority is None
                value = [];
                return
            end
            value = uiPriority.val;
        end
        function set.uiPriority_val(obj, value)
            obj.remove_uiPriority_();
            if ~isequal(value, [])   % Python: if value is not None
                uiPriority = obj.add_uiPriority_();
                uiPriority.val = value;
            end
        end

        % ---- unhideWhenUsed_val (styles.py 256-269) ----
        function value = get.unhideWhenUsed_val(obj)
            unhideWhenUsed = obj.unhideWhenUsed;
            if isequal(unhideWhenUsed, [])   % Python: if unhideWhenUsed is None
                value = false;               % Python: return False
                return
            end
            value = unhideWhenUsed.val;
        end
        function set.unhideWhenUsed_val(obj, value)
            obj.remove_unhideWhenUsed_();
            if ~isequal(value, []) && value   % Python: if bool(value) is True (H4)
                unhideWhenUsed = obj.add_unhideWhenUsed_();
                unhideWhenUsed.val = value;
            end
        end

        % ---- delete_style (styles.py 168-170) ----
        function delete_style(obj)
            % DELETE_STYLE Remove this `w:style` element from its parent `w:styles`.
            %
            %   Faithful port of CT_Style.delete (styles.py 168-170):
            %   self.getparent().remove(self).
            %
            %   Named `delete_style` (not `delete`) so this element method is NEVER
            %   MATLAB's handle destructor -- H17 dissolved: with no `delete`
            %   override, GC never calls this and no guarded-destructor machinery
            %   (try/catch, isvalid guard) is needed. The parent-guard keeps an
            %   explicit call on an UNPARENTED element a no-op, preserving the
            %   pre-rename behavior (python-docx would raise AttributeError there,
            %   but no python-docx caller reaches it).
            %
            %   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::CT_Style.delete
            p = obj.getparent();
            if ~isequal(p, [])   % Python: self.getparent().remove(self)
                p.remove(obj);
            end
        end
    end
end
