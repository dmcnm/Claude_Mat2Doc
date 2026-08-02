classdef CT_ShapeProperties < mat2doc.oxml.BaseOxmlElement
% CT_SHAPEPROPERTIES Custom element class for the <pic:spPr> element.
%
%   `<pic:spPr>` element, specifies size and shape of picture container. A single
%   ZeroOrOne descriptor `xfrm` ("a:xfrm", with successors) plus read/write
%   `cx` / `cy` @property members that route the shape width/height through the
%   a:xfrm/a:ext extent.
%
%   DOCX vs pptx: python-docx's CT_ShapeProperties is this MINIMAL class (xfrm +
%   cx/cy). python-pptx's same-named class (oxml/shapes/shared.py) is far richer
%   (custGeom/prstGeom, the EG_FillProperties choice group, ln/effectLst, a
%   gradFill override, xpath-based x/y/cx/cy). Ported to match DOCX (xfrm + cx/cy
%   only). Registered for <pic:spPr> in docx (oxml/__init__.py:57).
%
%   cx / cy (shape.py:239-263): READ returns Emu (via xfrm.cx / xfrm.cy) or []
%   (None) when a:xfrm is absent (H3); WRITE calls get_or_add_xfrm() then sets
%   xfrm.cx / xfrm.cy (which routes on into a:ext). This is the setter
%   CT_Picture.new uses to stamp the picture's display size.
%
%   UNDERSCORE ROTATION (design.md section 2): _new_xfrm -> new_xfrm_,
%   _insert_xfrm -> insert_xfrm_.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_ShapeProperties
%   (lines 222-263; registered for <pic:spPr>, oxml/__init__.py:57)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        XFRM_TAG = "a:xfrm"  % ZeroOrOne @ shape.py:225
        XFRM_SUCCESSORS = ["a:custGeom", "a:prstGeom", "a:ln", "a:effectLst", ...
            "a:effectDag", "a:scene3d", "a:sp3d", "a:extLst"]  % successors @ shape.py:227-236
    end

    properties (Dependent)  % generated descriptor properties + @property members
        xfrm  % ZeroOrOne('a:xfrm') -- child or [] (None)
        cx    % @property r/w: shape width via a:xfrm/a:ext, Emu or [] (None) if xfrm absent
        cy    % @property r/w: shape height via a:xfrm/a:ext, Emu or [] (None) if xfrm absent
    end

    methods
        function obj = CT_ShapeProperties(varargin)
            % CT_SHAPEPROPERTIES Construct a loose <pic:spPr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- xfrm (ZeroOrOne) ----
        function child = get.xfrm(obj)
            child = obj.getChild(obj.XFRM_TAG);
        end
        function child = get_or_add_xfrm(obj)
            child = obj.getOrAddChild(obj.XFRM_TAG, obj.XFRM_SUCCESSORS);
        end
        function child = new_xfrm_(obj)
            child = obj.newChild(obj.XFRM_TAG);
        end
        function child = insert_xfrm_(obj, child)
            child = obj.insertChildInSequence(child, obj.XFRM_SUCCESSORS);
        end
        function child = add_xfrm_(obj, varargin)
            child = obj.addChild(obj.XFRM_TAG, obj.XFRM_SUCCESSORS, varargin{:});
        end
        function remove_xfrm_(obj)
            obj.removeChild(obj.XFRM_TAG);
        end

        % ---- cx (@property r/w, shape.py:239-250) ----
        function value = get.cx(obj)
            % Python: xfrm = self.xfrm; if xfrm is None: return None; return xfrm.cx
            xfrm = obj.xfrm;
            if isequal(xfrm, [])   % Python: if xfrm is None (H3)
                value = [];
                return
            end
            value = xfrm.cx;
        end
        function set.cx(obj, value)
            % Python: xfrm = self.get_or_add_xfrm(); xfrm.cx = value
            xfrm = obj.get_or_add_xfrm();
            xfrm.cx = value;
        end

        % ---- cy (@property r/w, shape.py:252-263) ----
        function value = get.cy(obj)
            % Python: xfrm = self.xfrm; if xfrm is None: return None; return xfrm.cy
            xfrm = obj.xfrm;
            if isequal(xfrm, [])   % Python: if xfrm is None (H3)
                value = [];
                return
            end
            value = xfrm.cy;
        end
        function set.cy(obj, value)
            % Python: xfrm = self.get_or_add_xfrm(); xfrm.cy = value
            xfrm = obj.get_or_add_xfrm();
            xfrm.cy = value;
        end
    end
end
