classdef CT_Transform2D < mat2doc.oxml.BaseOxmlElement
% CT_TRANSFORM2D Custom element class for the <a:xfrm> element.
%
%   `<a:xfrm>` element, specifies size and shape of picture container. Two
%   ZeroOrOne descriptors `off` ("a:off", successors ("a:ext",)) and `ext`
%   ("a:ext", successors ()), plus read/write `cx` / `cy` @property members that
%   route width/height through the a:ext extent.
%
%   DOCX vs pptx: python-docx's CT_Transform2D has ONLY off / ext + cx / cy, with
%   NO _new_off/_new_ext override and NO rot/flipH/flipV/chOff/chExt/x/y surface.
%   python-pptx's same-named class (oxml/shapes/shared.py) adds chOff/chExt, the
%   rot/flipH/flipV attributes, x/y properties, and the _new_off/_new_ext
%   overrides (which pre-populate the RequiredAttribute children). Ported to
%   match DOCX -- off / ext use the GENERIC engine (no override), so a fresh
%   a:off / a:ext is created bare; docx's callers (CT_Picture.new) only ever set
%   cx/cy on a template that ALREADY carries a fully-formed a:off / a:ext, so the
%   generic path is never exercised on a bare create in the ported code paths.
%
%   cx / cy (shape.py:277-299): READ returns ext.cx / ext.cy or [] (None) when
%   a:ext is absent (H3); WRITE calls get_or_add_ext() then sets ext.cx / ext.cy.
%
%   UNDERSCORE ROTATION (design.md section 2): _new_off -> new_off_, _insert_off
%   -> insert_off_, _add_off -> add_off_ (and ext).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Transform2D
%   (lines 271-299; registered for <a:xfrm>, oxml/__init__.py:52)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NO_SUCCESSORS = string.empty(1, 0)   % ext: successors=() -> append at end
        OFF_TAG = "a:off"                    % ZeroOrOne @ shape.py:274
        OFF_SUCCESSORS = "a:ext"             % successors=("a:ext",) @ shape.py:274
        EXT_TAG = "a:ext"                    % ZeroOrOne @ shape.py:275
    end

    properties (Dependent)  % generated descriptor properties + @property members
        off  % ZeroOrOne('a:off') -- child or [] (None)
        ext  % ZeroOrOne('a:ext') -- child or [] (None)
        cx   % @property r/w: a:ext/@cx as Emu, or [] (None) if ext absent
        cy   % @property r/w: a:ext/@cy as Emu, or [] (None) if ext absent
    end

    methods
        function obj = CT_Transform2D(varargin)
            % CT_TRANSFORM2D Construct a loose <a:xfrm> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- off (ZeroOrOne) ----
        function child = get.off(obj)
            child = obj.getChild(obj.OFF_TAG);
        end
        function child = get_or_add_off(obj)
            child = obj.getOrAddChild(obj.OFF_TAG, obj.OFF_SUCCESSORS);
        end
        function child = new_off_(obj)
            child = obj.newChild(obj.OFF_TAG);
        end
        function child = insert_off_(obj, child)
            child = obj.insertChildInSequence(child, obj.OFF_SUCCESSORS);
        end
        function child = add_off_(obj, varargin)
            child = obj.addChild(obj.OFF_TAG, obj.OFF_SUCCESSORS, varargin{:});
        end
        function remove_off_(obj)
            obj.removeChild(obj.OFF_TAG);
        end

        % ---- ext (ZeroOrOne) ----
        function child = get.ext(obj)
            child = obj.getChild(obj.EXT_TAG);
        end
        function child = get_or_add_ext(obj)
            child = obj.getOrAddChild(obj.EXT_TAG, obj.NO_SUCCESSORS);
        end
        function child = new_ext_(obj)
            child = obj.newChild(obj.EXT_TAG);
        end
        function child = insert_ext_(obj, child)
            child = obj.insertChildInSequence(child, obj.NO_SUCCESSORS);
        end
        function child = add_ext_(obj, varargin)
            child = obj.addChild(obj.EXT_TAG, obj.NO_SUCCESSORS, varargin{:});
        end
        function remove_ext_(obj)
            obj.removeChild(obj.EXT_TAG);
        end

        % ---- cx (@property r/w, shape.py:277-287) ----
        function value = get.cx(obj)
            % Python: ext = self.ext; if ext is None: return None; return ext.cx
            ext = obj.ext;
            if isequal(ext, [])   % Python: if ext is None (H3)
                value = [];
                return
            end
            value = ext.cx;
        end
        function set.cx(obj, value)
            % Python: ext = self.get_or_add_ext(); ext.cx = value
            ext = obj.get_or_add_ext();
            ext.cx = value;
        end

        % ---- cy (@property r/w, shape.py:289-299) ----
        function value = get.cy(obj)
            % Python: ext = self.ext; if ext is None: return None; return ext.cy
            ext = obj.ext;
            if isequal(ext, [])   % Python: if ext is None (H3)
                value = [];
                return
            end
            value = ext.cy;
        end
        function set.cy(obj, value)
            % Python: ext = self.get_or_add_ext(); ext.cy = value
            ext = obj.get_or_add_ext();
            ext.cy = value;
        end
    end
end
