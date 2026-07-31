classdef CT_Styles < mat2doc.oxml.BaseOxmlElement
% CT_STYLES `<w:styles>` element, the root element of a styles part (styles.xml).
%
%   Contains an optional `<w:docDefaults>`, an optional `<w:latentStyles>`, and
%   zero-or-more `<w:style>` children. Provides the style-lookup surface the
%   Styles API (P4-7a) reads: get_by_id / get_by_name / default_for, plus the
%   builder add_style_of_type.
%
%   H11 (child ordering): _tag_seq (styles.py 275, VERBATIM) = (w:docDefaults,
%   w:latentStyles, w:style) is stored as TAG_SEQ. latentStyles = ZeroOrOne with
%   successors=_tag_seq[2:] -> TAG_SEQ(3:end) = ["w:style"] (H1: 0-based start 2
%   -> 1-based start 3). style = ZeroOrMore with successors=() -> APPEND. docx
%   ZeroOrMore also generates the PUBLIC add_style (D-delta-4), which
%   add_style_of_type calls.
%
%   XPATH forms (both in the mini-XPath verified subset, design.md section 3;
%   the second is a documented evaluate_xpath example):
%     get_by_id:   w:style[@w:styleId="<id>"]      (attribute-equality predicate)
%     get_by_name: w:style[w:name/@w:val="<name>"] (predicate attribute sub-path)
%   Both use next(iter(xpath), None) -> res(1) (H1, 1-based first) or [] (H3).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the single <w:styles> root of styles.xml.
%
%   Example:
%       styles = mat2doc.oxml.OxmlElement("w:styles");
%       s = styles.add_style_of_type("Heading 1", ...
%           mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH, true);
%       styles.get_by_id("Heading1") == s        % true (H5 identity)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::CT_Styles
%   (lines 272-320; registered for w:styles)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (styles.py 275; 3 tags)
        TAG_SEQ       = ["w:docDefaults", "w:latentStyles", "w:style"]
        LATENTSTYLES_TAG = "w:latentStyles"    % ZeroOrOne @ styles.py:276
        STYLE_TAG     = "w:style"              % ZeroOrMore @ styles.py:277
        NO_SUCCESSORS = string.empty(1, 0)     % successors=() -> APPEND
    end

    properties (Dependent)  % ZeroOrOne getter + ZeroOrMore list getter
        latentStyles  % <w:latentStyles> child or [] (None) if absent
        style_lst     % list of <w:style> children (document order)
    end

    methods
        function obj = CT_Styles(varargin)
            % CT_STYLES Construct a loose <w:styles> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- latentStyles (ZeroOrOne, successors=_tag_seq[2:] -> TAG_SEQ(3:end)) ----
        function child = get.latentStyles(obj);            child = obj.getChild(obj.LATENTSTYLES_TAG); end
        function child = get_or_add_latentStyles(obj);     child = obj.getOrAddChild(obj.LATENTSTYLES_TAG, obj.TAG_SEQ(3:end)); end
        function child = new_latentStyles_(obj);           child = obj.newChild(obj.LATENTSTYLES_TAG); end
        function child = insert_latentStyles_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(3:end)); end
        function child = add_latentStyles_(obj, varargin); child = obj.addChild(obj.LATENTSTYLES_TAG, obj.TAG_SEQ(3:end), varargin{:}); end
        function remove_latentStyles_(obj);                obj.removeChild(obj.LATENTSTYLES_TAG); end

        % ---- style (ZeroOrMore, successors=() -> APPEND) ----
        function lst = get.style_lst(obj);          lst = obj.getChildList(obj.STYLE_TAG); end
        function child = new_style_(obj);           child = obj.newChild(obj.STYLE_TAG); end
        function child = insert_style_(obj, child); child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS); end
        function child = add_style_(obj, varargin); child = obj.addChild(obj.STYLE_TAG, obj.NO_SUCCESSORS, varargin{:}); end
        function child = add_style(obj);            child = obj.add_style_(); end   % public adder (D-delta-4)

        % =================== methods (styles.py 280-320) ===================

        function style = add_style_of_type(obj, name, style_type, builtin)
            % ADD_STYLE_OF_TYPE Return a newly added `w:style` having `name` and `style_type`.
            %   `w:style/@customStyle` is set based on `builtin`.
            %   Ported from python-docx v1.2.0: styles.py CT_Styles.add_style_of_type
            %   (lines 280-290).
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_Styles
                name (1,1) string
                style_type (1,1) mat2doc.enum.style.WD_STYLE_TYPE
                builtin (1,1) logical
            end
            style = obj.add_style();
            style.type = style_type;
            if builtin
                style.customStyle = [];     % Python: None if builtin
            else
                style.customStyle = true;   % Python: else True
            end
            style.styleId = mat2doc.oxml.styles.styleId_from_name(name);
            style.name_val = name;
        end

        function value = default_for(obj, style_type)
            % DEFAULT_FOR Last `w:style[@w:type="<style_type>"][@w:default]` in document order, or [].
            %   Ported from python-docx v1.2.0: styles.py CT_Styles.default_for
            %   (lines 292-300). H4: `s.type == style_type and s.default` ->
            %   isequal(s.type, style_type) (None-safe enum compare) &&
            %   ~isequal(d, []) && d (None/False/True truthiness). Spec: LAST
            %   default in document order -> matches{end}.
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_Styles
                style_type (1,1) mat2doc.enum.style.WD_STYLE_TYPE
            end
            styles = obj.iter_styles_();
            matches = {};   % Python: default_styles_for_type = [...]
            for i = 1:numel(styles)
                s = styles(i);
                d = s.default;
                if isequal(s.type, style_type) && ~isequal(d, []) && d
                    matches{end + 1} = s; %#ok<AGROW>
                end
            end
            if isempty(matches)   % Python: if not default_styles_for_type
                value = [];
                return
            end
            value = matches{end};   % Python: return default_styles_for_type[-1]
        end

        function value = get_by_id(obj, styleId)
            % GET_BY_ID `w:style` child where @styleId = `styleId`, or [] if not found.
            %   Ported from python-docx v1.2.0: styles.py CT_Styles.get_by_id
            %   (lines 302-308): next(iter(self.xpath(f'w:style[@w:styleId=
            %   "{styleId}"]')), None). H1: res(1) is the first match; H3: [].
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_Styles
                styleId (1,1) string
            end
            res = obj.xpath("w:style[@w:styleId=""" + styleId + """]");
            if isempty(res)
                value = [];
                return
            end
            value = res(1);
        end

        function value = get_by_name(obj, name)
            % GET_BY_NAME `w:style` child with a `w:name` grandchild valued `name`, or [].
            %   Ported from python-docx v1.2.0: styles.py CT_Styles.get_by_name
            %   (lines 310-316): next(iter(self.xpath('w:style[w:name/@w:val="%s"]'
            %   % name)), None). H1: res(1); H3: [].
            arguments
                obj (1,1) mat2doc.oxml.styles.CT_Styles
                name (1,1) string
            end
            res = obj.xpath("w:style[w:name/@w:val=""" + name + """]");
            if isempty(res)
                value = [];
                return
            end
            value = res(1);
        end
    end

    methods (Access = private)
        function value = iter_styles_(obj)
            % ITER_STYLES_ Each `w:style` child element in document order.
            %   Ported from python-docx v1.2.0: styles.py CT_Styles._iter_styles
            %   (lines 318-320): (style for style in self.xpath("w:style")). H9:
            %   the lazy generator is materialized to the xpath array (default_for,
            %   the only caller, does not mutate during iteration).
            value = obj.xpath("w:style");
        end
    end
end
