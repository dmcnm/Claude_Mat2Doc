classdef CT_RPr < mat2doc.oxml.BaseOxmlElement
% CT_RPR `<w:rPr>` element: the run-properties container for a run.
%
%   THE M2 BYTE-CRITICAL element. Its 27 ZeroOrOne child descriptors carry
%   NON-CONTIGUOUS H11 successor slices of a single 39-element _tag_seq; the
%   slices drive insert_element_before, which re-sorts scrambled child adds
%   into canonical OOXML schema order so document.xml/styles.xml stay
%   byte-identical. (font.py's header calls _tag_seq a 40-tuple; the actual
%   tuple has 39 entries, font.py lines 80-118 -- see audit VERIFY-note. The
%   slices resolve identically either way; specVanish=[38:] -> (w:oMath,),
%   oMath=[39:] -> ().)
%
%   _tag_seq (font.py 79-119, VERBATIM, 39 tags) is stored as the Constant
%   TAG_SEQ. Each descriptor's successors = Python _tag_seq[s0:] is expressed
%   as TAG_SEQ(s0+1:end) (H1: 0-based Python slice start s0 -> 1-based MATLAB
%   start s0+1). The own tag of a descriptor sits at TAG_SEQ(s0) (1-based).
%
%   H11 SUCCESSOR-SLICE TABLE (prop : Python slice : MATLAB slice):
%     rStyle     = ZeroOrOne("w:rStyle", successors=_tag_seq[1:]) -> TAG_SEQ(2:end)
%     rFonts     = ZeroOrOne("w:rFonts", successors=_tag_seq[2:]) -> TAG_SEQ(3:end)
%     b          = ZeroOrOne("w:b", successors=_tag_seq[3:]) -> TAG_SEQ(4:end)
%     bCs        = ZeroOrOne("w:bCs", successors=_tag_seq[4:]) -> TAG_SEQ(5:end)
%     i          = ZeroOrOne("w:i", successors=_tag_seq[5:]) -> TAG_SEQ(6:end)
%     iCs        = ZeroOrOne("w:iCs", successors=_tag_seq[6:]) -> TAG_SEQ(7:end)
%     caps       = ZeroOrOne("w:caps", successors=_tag_seq[7:]) -> TAG_SEQ(8:end)
%     smallCaps  = ZeroOrOne("w:smallCaps", successors=_tag_seq[8:]) -> TAG_SEQ(9:end)
%     strike     = ZeroOrOne("w:strike", successors=_tag_seq[9:]) -> TAG_SEQ(10:end)
%     dstrike    = ZeroOrOne("w:dstrike", successors=_tag_seq[10:]) -> TAG_SEQ(11:end)
%     outline    = ZeroOrOne("w:outline", successors=_tag_seq[11:]) -> TAG_SEQ(12:end)
%     shadow     = ZeroOrOne("w:shadow", successors=_tag_seq[12:]) -> TAG_SEQ(13:end)
%     emboss     = ZeroOrOne("w:emboss", successors=_tag_seq[13:]) -> TAG_SEQ(14:end)
%     imprint    = ZeroOrOne("w:imprint", successors=_tag_seq[14:]) -> TAG_SEQ(15:end)
%     noProof    = ZeroOrOne("w:noProof", successors=_tag_seq[15:]) -> TAG_SEQ(16:end)
%     snapToGrid = ZeroOrOne("w:snapToGrid", successors=_tag_seq[16:]) -> TAG_SEQ(17:end)
%     vanish     = ZeroOrOne("w:vanish", successors=_tag_seq[17:]) -> TAG_SEQ(18:end)
%     webHidden  = ZeroOrOne("w:webHidden", successors=_tag_seq[18:]) -> TAG_SEQ(19:end)
%     color      = ZeroOrOne("w:color", successors=_tag_seq[19:]) -> TAG_SEQ(20:end)
%     sz         = ZeroOrOne("w:sz", successors=_tag_seq[24:]) -> TAG_SEQ(25:end)
%     highlight  = ZeroOrOne("w:highlight", successors=_tag_seq[26:]) -> TAG_SEQ(27:end)
%     u          = ZeroOrOne("w:u", successors=_tag_seq[27:]) -> TAG_SEQ(28:end)
%     vertAlign  = ZeroOrOne("w:vertAlign", successors=_tag_seq[32:]) -> TAG_SEQ(33:end)
%     rtl        = ZeroOrOne("w:rtl", successors=_tag_seq[33:]) -> TAG_SEQ(34:end)
%     cs         = ZeroOrOne("w:cs", successors=_tag_seq[34:]) -> TAG_SEQ(35:end)
%     specVanish = ZeroOrOne("w:specVanish", successors=_tag_seq[38:]) -> TAG_SEQ(39:end)
%     oMath      = ZeroOrOne("w:oMath", successors=_tag_seq[39:]) -> TAG_SEQ(40:end)
%
%   _new_color OVERRIDE (font.py 149-151): CT_RPr defines `_new_color`, which
%   the metaclass lets WIN over the generated default creator, so
%   get_or_add_color / add_color_ route through new_color_ (seeds
%   <w:color w:val="000000"/>, RGB black) rather than the generic engine
%   creator. Ported per the design.md section-2 OVERRIDE alert (same shape as
%   Mat2Ppt CT_Chart._new_title). All OTHER 26 descriptors use the generic
%   BaseOxmlElement engine (getChild/getOrAddChild/newChild/addChild/
%   insertChildInSequence/removeChild).
%
%   GENERATED DESCRIPTOR FAMILY (per ZeroOrOne, xmlchemy docx form): get.x,
%   get_or_add_x, new_x_, insert_x_, add_x_, remove_x_ (underscore rotation:
%   Python _new_x/_insert_x/_add_x/_remove_x -> new_x_/insert_x_/add_x_/
%   remove_x_; get_or_add_x is public). The pyright callable annotations in
%   font.py 64-77 (get_or_add_color/_add_rStyle/_remove_u/...) are type hints
%   only and add no members beyond this family.
%
%   @property members (font.py 153-305): highlight_val, rFonts_ascii,
%   rFonts_hAnsi, style, subscript, superscript, sz_val, u_val -- all ported
%   LIVE (deps all present). NOTE the rFonts_hAnsi setter is ASYMMETRIC vs
%   rFonts_ascii: it does NOT remove w:rFonts on None, it sets hAnsi=None
%   (removing only @w:hAnsi) -- ported verbatim (font.py 201-206). The
%   _get_bool_val/_set_bool_val helpers (font.py 307-319) port to
%   get_bool_val_/set_bool_val_ (used by the Font proxy, a later WP).
%
%   H3 (None): inline isequal(x, []) (established Mat2Doc oxml convention).
%   H4 (truthiness): the subscript/superscript setter middle branch
%   `bool(value) is True` -> `elseif value` (value is logical/None here).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the
%   parser instantiates this on the many <w:rPr> nodes inside a real
%   styles.xml/document.xml.
%
%   Example:
%       rPr = mat2doc.oxml.OxmlElement("w:rPr");
%       rPr.get_or_add_u().val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
%       rPr.get_or_add_b();          % adds <w:b/>, re-sorted before <w:u> (H11)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/font.py::CT_RPr
%   (lines 61-319; registered for w:rPr)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (font.py 79-119; 39 tags)
        TAG_SEQ = [ ...
            "w:rStyle", "w:rFonts", "w:b", "w:bCs", ...
            "w:i", "w:iCs", "w:caps", "w:smallCaps", ...
            "w:strike", "w:dstrike", "w:outline", "w:shadow", ...
            "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid", ...
            "w:vanish", "w:webHidden", "w:color", "w:spacing", ...
            "w:w", "w:kern", "w:position", "w:sz", ...
            "w:szCs", "w:highlight", "w:u", "w:effect", ...
            "w:bdr", "w:shd", "w:fitText", "w:vertAlign", ...
            "w:rtl", "w:cs", "w:em", "w:lang", ...
            "w:eastAsianLayout", "w:specVanish", "w:oMath" ]
    end

    properties (Dependent)  % generated ZeroOrOne getters + @property members
        % -- 27 ZeroOrOne child getters (read-only; use get_or_add_x/remove_x_) --
        rStyle
        rFonts
        b
        bCs
        i
        iCs
        caps
        smallCaps
        strike
        dstrike
        outline
        shadow
        emboss
        imprint
        noProof
        snapToGrid
        vanish
        webHidden
        color
        sz
        highlight
        u
        vertAlign
        rtl
        cs
        specVanish
        oMath
        % -- @property members (font.py 153-305) --
        highlight_val   % ./w:highlight/@val or []
        rFonts_ascii    % w:rFonts/@w:ascii or []
        rFonts_hAnsi    % w:rFonts/@w:hAnsi or []
        style           % ./w:rStyle/@val or []
        subscript       % True/False/[] per ./w:vertAlign/@w:val
        superscript     % True/False/[] per ./w:vertAlign/@w:val
        sz_val          % w:sz/@w:val (Length) or []
        u_val           % w:u/@val (WD_UNDERLINE) or []
    end

    methods
        function obj = CT_RPr(varargin)
            % CT_RPR Construct a loose <w:rPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- rStyle (ZeroOrOne, successors=_tag_seq[1:] -> TAG_SEQ(2:end)) ----
        function child = get.rStyle(obj);            child = obj.getChild("w:rStyle"); end
        function child = get_or_add_rStyle(obj);     child = obj.getOrAddChild("w:rStyle", obj.TAG_SEQ(2:end)); end
        function child = new_rStyle_(obj);           child = obj.newChild("w:rStyle"); end
        function child = insert_rStyle_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(2:end)); end
        function child = add_rStyle_(obj, varargin); child = obj.addChild("w:rStyle", obj.TAG_SEQ(2:end), varargin{:}); end
        function remove_rStyle_(obj);                obj.removeChild("w:rStyle"); end

        % ---- rFonts (ZeroOrOne, successors=_tag_seq[2:] -> TAG_SEQ(3:end)) ----
        function child = get.rFonts(obj);            child = obj.getChild("w:rFonts"); end
        function child = get_or_add_rFonts(obj);     child = obj.getOrAddChild("w:rFonts", obj.TAG_SEQ(3:end)); end
        function child = new_rFonts_(obj);           child = obj.newChild("w:rFonts"); end
        function child = insert_rFonts_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(3:end)); end
        function child = add_rFonts_(obj, varargin); child = obj.addChild("w:rFonts", obj.TAG_SEQ(3:end), varargin{:}); end
        function remove_rFonts_(obj);                obj.removeChild("w:rFonts"); end

        % ---- b (ZeroOrOne, successors=_tag_seq[3:] -> TAG_SEQ(4:end)) ----
        function child = get.b(obj);            child = obj.getChild("w:b"); end
        function child = get_or_add_b(obj);     child = obj.getOrAddChild("w:b", obj.TAG_SEQ(4:end)); end
        function child = new_b_(obj);           child = obj.newChild("w:b"); end
        function child = insert_b_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(4:end)); end
        function child = add_b_(obj, varargin); child = obj.addChild("w:b", obj.TAG_SEQ(4:end), varargin{:}); end
        function remove_b_(obj);                obj.removeChild("w:b"); end

        % ---- bCs (ZeroOrOne, successors=_tag_seq[4:] -> TAG_SEQ(5:end)) ----
        function child = get.bCs(obj);            child = obj.getChild("w:bCs"); end
        function child = get_or_add_bCs(obj);     child = obj.getOrAddChild("w:bCs", obj.TAG_SEQ(5:end)); end
        function child = new_bCs_(obj);           child = obj.newChild("w:bCs"); end
        function child = insert_bCs_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(5:end)); end
        function child = add_bCs_(obj, varargin); child = obj.addChild("w:bCs", obj.TAG_SEQ(5:end), varargin{:}); end
        function remove_bCs_(obj);                obj.removeChild("w:bCs"); end

        % ---- i (ZeroOrOne, successors=_tag_seq[5:] -> TAG_SEQ(6:end)) ----
        function child = get.i(obj);            child = obj.getChild("w:i"); end
        function child = get_or_add_i(obj);     child = obj.getOrAddChild("w:i", obj.TAG_SEQ(6:end)); end
        function child = new_i_(obj);           child = obj.newChild("w:i"); end
        function child = insert_i_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(6:end)); end
        function child = add_i_(obj, varargin); child = obj.addChild("w:i", obj.TAG_SEQ(6:end), varargin{:}); end
        function remove_i_(obj);                obj.removeChild("w:i"); end

        % ---- iCs (ZeroOrOne, successors=_tag_seq[6:] -> TAG_SEQ(7:end)) ----
        function child = get.iCs(obj);            child = obj.getChild("w:iCs"); end
        function child = get_or_add_iCs(obj);     child = obj.getOrAddChild("w:iCs", obj.TAG_SEQ(7:end)); end
        function child = new_iCs_(obj);           child = obj.newChild("w:iCs"); end
        function child = insert_iCs_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(7:end)); end
        function child = add_iCs_(obj, varargin); child = obj.addChild("w:iCs", obj.TAG_SEQ(7:end), varargin{:}); end
        function remove_iCs_(obj);                obj.removeChild("w:iCs"); end

        % ---- caps (ZeroOrOne, successors=_tag_seq[7:] -> TAG_SEQ(8:end)) ----
        function child = get.caps(obj);            child = obj.getChild("w:caps"); end
        function child = get_or_add_caps(obj);     child = obj.getOrAddChild("w:caps", obj.TAG_SEQ(8:end)); end
        function child = new_caps_(obj);           child = obj.newChild("w:caps"); end
        function child = insert_caps_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(8:end)); end
        function child = add_caps_(obj, varargin); child = obj.addChild("w:caps", obj.TAG_SEQ(8:end), varargin{:}); end
        function remove_caps_(obj);                obj.removeChild("w:caps"); end

        % ---- smallCaps (ZeroOrOne, successors=_tag_seq[8:] -> TAG_SEQ(9:end)) ----
        function child = get.smallCaps(obj);            child = obj.getChild("w:smallCaps"); end
        function child = get_or_add_smallCaps(obj);     child = obj.getOrAddChild("w:smallCaps", obj.TAG_SEQ(9:end)); end
        function child = new_smallCaps_(obj);           child = obj.newChild("w:smallCaps"); end
        function child = insert_smallCaps_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(9:end)); end
        function child = add_smallCaps_(obj, varargin); child = obj.addChild("w:smallCaps", obj.TAG_SEQ(9:end), varargin{:}); end
        function remove_smallCaps_(obj);                obj.removeChild("w:smallCaps"); end

        % ---- strike (ZeroOrOne, successors=_tag_seq[9:] -> TAG_SEQ(10:end)) ----
        function child = get.strike(obj);            child = obj.getChild("w:strike"); end
        function child = get_or_add_strike(obj);     child = obj.getOrAddChild("w:strike", obj.TAG_SEQ(10:end)); end
        function child = new_strike_(obj);           child = obj.newChild("w:strike"); end
        function child = insert_strike_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(10:end)); end
        function child = add_strike_(obj, varargin); child = obj.addChild("w:strike", obj.TAG_SEQ(10:end), varargin{:}); end
        function remove_strike_(obj);                obj.removeChild("w:strike"); end

        % ---- dstrike (ZeroOrOne, successors=_tag_seq[10:] -> TAG_SEQ(11:end)) ----
        function child = get.dstrike(obj);            child = obj.getChild("w:dstrike"); end
        function child = get_or_add_dstrike(obj);     child = obj.getOrAddChild("w:dstrike", obj.TAG_SEQ(11:end)); end
        function child = new_dstrike_(obj);           child = obj.newChild("w:dstrike"); end
        function child = insert_dstrike_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(11:end)); end
        function child = add_dstrike_(obj, varargin); child = obj.addChild("w:dstrike", obj.TAG_SEQ(11:end), varargin{:}); end
        function remove_dstrike_(obj);                obj.removeChild("w:dstrike"); end

        % ---- outline (ZeroOrOne, successors=_tag_seq[11:] -> TAG_SEQ(12:end)) ----
        function child = get.outline(obj);            child = obj.getChild("w:outline"); end
        function child = get_or_add_outline(obj);     child = obj.getOrAddChild("w:outline", obj.TAG_SEQ(12:end)); end
        function child = new_outline_(obj);           child = obj.newChild("w:outline"); end
        function child = insert_outline_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(12:end)); end
        function child = add_outline_(obj, varargin); child = obj.addChild("w:outline", obj.TAG_SEQ(12:end), varargin{:}); end
        function remove_outline_(obj);                obj.removeChild("w:outline"); end

        % ---- shadow (ZeroOrOne, successors=_tag_seq[12:] -> TAG_SEQ(13:end)) ----
        function child = get.shadow(obj);            child = obj.getChild("w:shadow"); end
        function child = get_or_add_shadow(obj);     child = obj.getOrAddChild("w:shadow", obj.TAG_SEQ(13:end)); end
        function child = new_shadow_(obj);           child = obj.newChild("w:shadow"); end
        function child = insert_shadow_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(13:end)); end
        function child = add_shadow_(obj, varargin); child = obj.addChild("w:shadow", obj.TAG_SEQ(13:end), varargin{:}); end
        function remove_shadow_(obj);                obj.removeChild("w:shadow"); end

        % ---- emboss (ZeroOrOne, successors=_tag_seq[13:] -> TAG_SEQ(14:end)) ----
        function child = get.emboss(obj);            child = obj.getChild("w:emboss"); end
        function child = get_or_add_emboss(obj);     child = obj.getOrAddChild("w:emboss", obj.TAG_SEQ(14:end)); end
        function child = new_emboss_(obj);           child = obj.newChild("w:emboss"); end
        function child = insert_emboss_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(14:end)); end
        function child = add_emboss_(obj, varargin); child = obj.addChild("w:emboss", obj.TAG_SEQ(14:end), varargin{:}); end
        function remove_emboss_(obj);                obj.removeChild("w:emboss"); end

        % ---- imprint (ZeroOrOne, successors=_tag_seq[14:] -> TAG_SEQ(15:end)) ----
        function child = get.imprint(obj);            child = obj.getChild("w:imprint"); end
        function child = get_or_add_imprint(obj);     child = obj.getOrAddChild("w:imprint", obj.TAG_SEQ(15:end)); end
        function child = new_imprint_(obj);           child = obj.newChild("w:imprint"); end
        function child = insert_imprint_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(15:end)); end
        function child = add_imprint_(obj, varargin); child = obj.addChild("w:imprint", obj.TAG_SEQ(15:end), varargin{:}); end
        function remove_imprint_(obj);                obj.removeChild("w:imprint"); end

        % ---- noProof (ZeroOrOne, successors=_tag_seq[15:] -> TAG_SEQ(16:end)) ----
        function child = get.noProof(obj);            child = obj.getChild("w:noProof"); end
        function child = get_or_add_noProof(obj);     child = obj.getOrAddChild("w:noProof", obj.TAG_SEQ(16:end)); end
        function child = new_noProof_(obj);           child = obj.newChild("w:noProof"); end
        function child = insert_noProof_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(16:end)); end
        function child = add_noProof_(obj, varargin); child = obj.addChild("w:noProof", obj.TAG_SEQ(16:end), varargin{:}); end
        function remove_noProof_(obj);                obj.removeChild("w:noProof"); end

        % ---- snapToGrid (ZeroOrOne, successors=_tag_seq[16:] -> TAG_SEQ(17:end)) ----
        function child = get.snapToGrid(obj);            child = obj.getChild("w:snapToGrid"); end
        function child = get_or_add_snapToGrid(obj);     child = obj.getOrAddChild("w:snapToGrid", obj.TAG_SEQ(17:end)); end
        function child = new_snapToGrid_(obj);           child = obj.newChild("w:snapToGrid"); end
        function child = insert_snapToGrid_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(17:end)); end
        function child = add_snapToGrid_(obj, varargin); child = obj.addChild("w:snapToGrid", obj.TAG_SEQ(17:end), varargin{:}); end
        function remove_snapToGrid_(obj);                obj.removeChild("w:snapToGrid"); end

        % ---- vanish (ZeroOrOne, successors=_tag_seq[17:] -> TAG_SEQ(18:end)) ----
        function child = get.vanish(obj);            child = obj.getChild("w:vanish"); end
        function child = get_or_add_vanish(obj);     child = obj.getOrAddChild("w:vanish", obj.TAG_SEQ(18:end)); end
        function child = new_vanish_(obj);           child = obj.newChild("w:vanish"); end
        function child = insert_vanish_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(18:end)); end
        function child = add_vanish_(obj, varargin); child = obj.addChild("w:vanish", obj.TAG_SEQ(18:end), varargin{:}); end
        function remove_vanish_(obj);                obj.removeChild("w:vanish"); end

        % ---- webHidden (ZeroOrOne, successors=_tag_seq[18:] -> TAG_SEQ(19:end)) ----
        function child = get.webHidden(obj);            child = obj.getChild("w:webHidden"); end
        function child = get_or_add_webHidden(obj);     child = obj.getOrAddChild("w:webHidden", obj.TAG_SEQ(19:end)); end
        function child = new_webHidden_(obj);           child = obj.newChild("w:webHidden"); end
        function child = insert_webHidden_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(19:end)); end
        function child = add_webHidden_(obj, varargin); child = obj.addChild("w:webHidden", obj.TAG_SEQ(19:end), varargin{:}); end
        function remove_webHidden_(obj);                obj.removeChild("w:webHidden"); end

        % ---- color (ZeroOrOne, successors=_tag_seq[19:] -> TAG_SEQ(20:end)) ----
        %      _new_color OVERRIDE (font.py 149-151): get_or_add / add route
        %      through new_color_ (seeds @w:val=000000, RGB black), NOT the
        %      generic engine creator.
        function child = get.color(obj);         child = obj.getChild("w:color"); end
        function child = get_or_add_color(obj)   % routed through the override
            child = obj.color;
            if isequal(child, [])                % Python: if child is None (H3)
                child = obj.add_color_();
            end
        end
        function child = new_color_(obj) %#ok<MANU>   % Python: _new_color (font.py 149-151)
            % Python: return parse_xml('<w:color %s w:val="000000"/>' % nsdecls("w"))
            child = mat2doc.oxml.parse_xml( ...
                "<w:color " + mat2doc.oxml.nsdecls("w") + " w:val=""000000""/>");
        end
        function child = insert_color_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(20:end)); end
        function child = add_color_(obj, varargin)    % routed through the override (new_color_)
            child = obj.new_color_();
            for k = 1:2:numel(varargin)
                child.(varargin{k}) = varargin{k + 1};
            end
            child = obj.insert_color_(child);
        end
        function remove_color_(obj);             obj.removeChild("w:color"); end

        % ---- sz (ZeroOrOne, successors=_tag_seq[24:] -> TAG_SEQ(25:end)) ----
        function child = get.sz(obj);            child = obj.getChild("w:sz"); end
        function child = get_or_add_sz(obj);     child = obj.getOrAddChild("w:sz", obj.TAG_SEQ(25:end)); end
        function child = new_sz_(obj);           child = obj.newChild("w:sz"); end
        function child = insert_sz_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(25:end)); end
        function child = add_sz_(obj, varargin); child = obj.addChild("w:sz", obj.TAG_SEQ(25:end), varargin{:}); end
        function remove_sz_(obj);                obj.removeChild("w:sz"); end

        % ---- highlight (ZeroOrOne, successors=_tag_seq[26:] -> TAG_SEQ(27:end)) ----
        function child = get.highlight(obj);            child = obj.getChild("w:highlight"); end
        function child = get_or_add_highlight(obj);     child = obj.getOrAddChild("w:highlight", obj.TAG_SEQ(27:end)); end
        function child = new_highlight_(obj);           child = obj.newChild("w:highlight"); end
        function child = insert_highlight_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(27:end)); end
        function child = add_highlight_(obj, varargin); child = obj.addChild("w:highlight", obj.TAG_SEQ(27:end), varargin{:}); end
        function remove_highlight_(obj);                obj.removeChild("w:highlight"); end

        % ---- u (ZeroOrOne, successors=_tag_seq[27:] -> TAG_SEQ(28:end)) ----
        function child = get.u(obj);            child = obj.getChild("w:u"); end
        function child = get_or_add_u(obj);     child = obj.getOrAddChild("w:u", obj.TAG_SEQ(28:end)); end
        function child = new_u_(obj);           child = obj.newChild("w:u"); end
        function child = insert_u_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(28:end)); end
        function child = add_u_(obj, varargin); child = obj.addChild("w:u", obj.TAG_SEQ(28:end), varargin{:}); end
        function remove_u_(obj);                obj.removeChild("w:u"); end

        % ---- vertAlign (ZeroOrOne, successors=_tag_seq[32:] -> TAG_SEQ(33:end)) ----
        function child = get.vertAlign(obj);            child = obj.getChild("w:vertAlign"); end
        function child = get_or_add_vertAlign(obj);     child = obj.getOrAddChild("w:vertAlign", obj.TAG_SEQ(33:end)); end
        function child = new_vertAlign_(obj);           child = obj.newChild("w:vertAlign"); end
        function child = insert_vertAlign_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(33:end)); end
        function child = add_vertAlign_(obj, varargin); child = obj.addChild("w:vertAlign", obj.TAG_SEQ(33:end), varargin{:}); end
        function remove_vertAlign_(obj);                obj.removeChild("w:vertAlign"); end

        % ---- rtl (ZeroOrOne, successors=_tag_seq[33:] -> TAG_SEQ(34:end)) ----
        function child = get.rtl(obj);            child = obj.getChild("w:rtl"); end
        function child = get_or_add_rtl(obj);     child = obj.getOrAddChild("w:rtl", obj.TAG_SEQ(34:end)); end
        function child = new_rtl_(obj);           child = obj.newChild("w:rtl"); end
        function child = insert_rtl_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(34:end)); end
        function child = add_rtl_(obj, varargin); child = obj.addChild("w:rtl", obj.TAG_SEQ(34:end), varargin{:}); end
        function remove_rtl_(obj);                obj.removeChild("w:rtl"); end

        % ---- cs (ZeroOrOne, successors=_tag_seq[34:] -> TAG_SEQ(35:end)) ----
        function child = get.cs(obj);            child = obj.getChild("w:cs"); end
        function child = get_or_add_cs(obj);     child = obj.getOrAddChild("w:cs", obj.TAG_SEQ(35:end)); end
        function child = new_cs_(obj);           child = obj.newChild("w:cs"); end
        function child = insert_cs_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(35:end)); end
        function child = add_cs_(obj, varargin); child = obj.addChild("w:cs", obj.TAG_SEQ(35:end), varargin{:}); end
        function remove_cs_(obj);                obj.removeChild("w:cs"); end

        % ---- specVanish (ZeroOrOne, successors=_tag_seq[38:] -> TAG_SEQ(39:end)) ----
        function child = get.specVanish(obj);            child = obj.getChild("w:specVanish"); end
        function child = get_or_add_specVanish(obj);     child = obj.getOrAddChild("w:specVanish", obj.TAG_SEQ(39:end)); end
        function child = new_specVanish_(obj);           child = obj.newChild("w:specVanish"); end
        function child = insert_specVanish_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(39:end)); end
        function child = add_specVanish_(obj, varargin); child = obj.addChild("w:specVanish", obj.TAG_SEQ(39:end), varargin{:}); end
        function remove_specVanish_(obj);                obj.removeChild("w:specVanish"); end

        % ---- oMath (ZeroOrOne, successors=_tag_seq[39:] -> TAG_SEQ(40:end)) ----
        function child = get.oMath(obj);            child = obj.getChild("w:oMath"); end
        function child = get_or_add_oMath(obj);     child = obj.getOrAddChild("w:oMath", obj.TAG_SEQ(40:end)); end
        function child = new_oMath_(obj);           child = obj.newChild("w:oMath"); end
        function child = insert_oMath_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(40:end)); end
        function child = add_oMath_(obj, varargin); child = obj.addChild("w:oMath", obj.TAG_SEQ(40:end), varargin{:}); end
        function remove_oMath_(obj);                obj.removeChild("w:oMath"); end

        % =================== @property members (font.py 153-305) ===================

        % ---- highlight_val (font.py 153-170) ----
        function value = get.highlight_val(obj)
            highlight = obj.highlight;
            if isequal(highlight, [])   % Python: if highlight is None
                value = [];
                return
            end
            value = highlight.val;
        end
        function set.highlight_val(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_highlight_();
                return
            end
            highlight = obj.get_or_add_highlight();
            highlight.val = value;
        end

        % ---- rFonts_ascii (font.py 172-191) ----
        function value = get.rFonts_ascii(obj)
            rFonts = obj.rFonts;
            if isequal(rFonts, [])
                value = [];
                return
            end
            value = rFonts.ascii;
        end
        function set.rFonts_ascii(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_rFonts_();
                return
            end
            rFonts = obj.get_or_add_rFonts();
            rFonts.ascii = value;
        end

        % ---- rFonts_hAnsi (font.py 193-206) -- ASYMMETRIC setter (see header) ----
        function value = get.rFonts_hAnsi(obj)
            rFonts = obj.rFonts;
            if isequal(rFonts, [])
                value = [];
                return
            end
            value = rFonts.hAnsi;
        end
        function set.rFonts_hAnsi(obj, value)
            % Python (font.py 202-206): if value is None and self.rFonts is None: return
            if isequal(value, []) && isequal(obj.rFonts, [])
                return
            end
            rFonts = obj.get_or_add_rFonts();
            rFonts.hAnsi = value;   % value may be [] -> CT_Fonts.hAnsi setter removes @w:hAnsi
        end

        % ---- style (font.py 208-227) ----
        function value = get.style(obj)
            rStyle = obj.rStyle;
            if isequal(rStyle, [])
                value = [];
                return
            end
            value = rStyle.val;
        end
        function set.style(obj, style)
            if isequal(style, [])           % Python: if style is None
                obj.remove_rStyle_();
            elseif isequal(obj.rStyle, [])  % Python: elif self.rStyle is None
                obj.add_rStyle_("val", style);   % Python: self._add_rStyle(val=style)
            else
                obj.rStyle.val = style;
            end
        end

        % ---- subscript (font.py) ----
        function value = get.subscript(obj)
            vertAlign = obj.vertAlign;
            if isequal(vertAlign, [])   % Python: if vertAlign is None
                value = [];
                return
            end
            % Python: return vertAlign.val == ST_VerticalAlignRun.SUBSCRIPT
            value = vertAlign.val == mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUBSCRIPT;
        end
        function set.subscript(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_vertAlign_();
            elseif value                % Python: elif bool(value) is True (H4)
                va = obj.get_or_add_vertAlign();
                va.val = mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUBSCRIPT;
            % -- Python: assert bool(value) is False --
            elseif ~isequal(obj.vertAlign, []) && obj.vertAlign.val == mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUBSCRIPT
                obj.remove_vertAlign_();
            end
        end

        % ---- superscript (font.py) ----
        function value = get.superscript(obj)
            vertAlign = obj.vertAlign;
            if isequal(vertAlign, [])   % Python: if vertAlign is None
                value = [];
                return
            end
            % Python: return vertAlign.val == ST_VerticalAlignRun.SUPERSCRIPT
            value = vertAlign.val == mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUPERSCRIPT;
        end
        function set.superscript(obj, value)
            if isequal(value, [])       % Python: if value is None
                obj.remove_vertAlign_();
            elseif value                % Python: elif bool(value) is True (H4)
                va = obj.get_or_add_vertAlign();
                va.val = mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUPERSCRIPT;
            % -- Python: assert bool(value) is False --
            elseif ~isequal(obj.vertAlign, []) && obj.vertAlign.val == mat2doc.oxml.simpletypes.ST_VerticalAlignRun.SUPERSCRIPT
                obj.remove_vertAlign_();
            end
        end

        % ---- sz_val (font.py 273-287) ----
        function value = get.sz_val(obj)
            sz = obj.sz;
            if isequal(sz, [])
                value = [];
                return
            end
            value = sz.val;
        end
        function set.sz_val(obj, value)
            if isequal(value, [])
                obj.remove_sz_();
                return
            end
            sz = obj.get_or_add_sz();
            sz.val = value;
        end

        % ---- u_val (font.py 289-305) ----
        function value = get.u_val(obj)
            u = obj.u;
            if isequal(u, [])
                value = [];
                return
            end
            value = u.val;
        end
        function set.u_val(obj, value)
            obj.remove_u_();                 % Python: self._remove_u()
            if ~isequal(value, [])           % Python: if value is not None
                added = obj.add_u_();        % Python: self._add_u()
                added.val = value;
            end
        end

        % ---- get_bool_val_ / set_bool_val_ (font.py 307-319; Font-proxy helpers) ----
        function value = get_bool_val_(obj, name)
            % Python _get_bool_val (font.py 307-312). `name` is the descriptor
            % PROPERTY name ("b","i","smallCaps",...); obj.(name) is its getter.
            element = obj.(name);
            if isequal(element, [])   % Python: if element is None
                value = [];
                return
            end
            value = element.val;
        end
        function set_bool_val_(obj, name, value)
            % Python _set_bool_val (font.py 314-319). Rotation: Python
            % _remove_<name> -> remove_<name>_ (feval dispatch on obj).
            if isequal(value, [])   % Python: if value is None
                feval("remove_" + name + "_", obj);
                return
            end
            element = feval("get_or_add_" + name, obj);
            element.val = value;
        end
    end
end
