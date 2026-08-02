classdef InlineShape < handle
% INLINESHAPE Proxy for a <wp:inline> element (container for an inline graphical object).
%
%   Represents the container for an inline graphical object -- most commonly an
%   inline picture. In python-docx InlineShape is a PLAIN object (not Parented,
%   not ElementProxy): it holds only the wrapped <wp:inline> element (`_inline`)
%   and compares by default object identity. Ported as `classdef ... < handle`
%   with the rotated attribute `inline_` and no eq/ne override (MATLAB default
%   handle identity == Python default object identity).
%
%   ATTRIBUTES (shape.py 55-57): Python `self._inline = inline`. Ported: inline_.
%
%   height / width (shape.py 59-70, 92-103): read/write display size as an |Emu|
%   (Length) via wp:inline/wp:extent; the setter ALSO writes the picture's
%   pic:spPr extent (graphic/graphicData/pic/spPr) so the extent and the picture
%   frame stay in sync. This WP ports the READ + WRITE surface faithfully; the
%   setter's pic path requires an inline picture (a graphicData carrying a
%   pic:pic) -- exactly python-docx's own precondition.
%
%   type (shape.py 72-90, H10): the WD_INLINE_SHAPE member implied by the
%   a:graphicData @uri (and, for a picture, whether the a:blip carries an
%   r:link). Compares @uri against the fixed nsmap URIs for pic / c / dgm.
%
%   PROPERTY-AS-DEPENDENT (design.md section 2): Python @property -> Dependent
%   property with get./set.; `type` is read-only.
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_inline` -> inline_.
%
%   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShape (lines 51-104)

    properties (Access = private)
        inline_   % _inline (shape.py 57): the wrapped <wp:inline> (a CT_Inline)
    end

    properties (Dependent)
        height  % r/w: display height as an Emu (Length)
        width   % r/w: display width as an Emu (Length)
        type    % read-only: WD_INLINE_SHAPE member implied by graphicData @uri
    end

    methods
        function obj = InlineShape(inline)
            % INLINESHAPE Wrap a <wp:inline> element (shape.py 55-57).
            %   Python: super().__init__(); self._inline = inline.
            %
            %   Inputs:  inline - the wrapped <wp:inline> (a CT_Inline).
            %   Outputs: obj    - a scalar InlineShape handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shape.py::InlineShape.__init__
            obj.inline_ = inline;
        end

        % ---- height (@property r/w, shape.py 59-70) ----
        function value = get.height(obj)
            % Python: return self._inline.extent.cy
            value = obj.inline_.extent.cy;
        end
        function set.height(obj, cy)
            % Python:
            %   self._inline.extent.cy = cy
            %   self._inline.graphic.graphicData.pic.spPr.cy = cy
            obj.inline_.extent.cy = cy;
            obj.inline_.graphic.graphicData.pic.spPr.cy = cy;
        end

        % ---- width (@property r/w, shape.py 92-103) ----
        function value = get.width(obj)
            % Python: return self._inline.extent.cx
            value = obj.inline_.extent.cx;
        end
        function set.width(obj, cx)
            % Python:
            %   self._inline.extent.cx = cx
            %   self._inline.graphic.graphicData.pic.spPr.cx = cx
            obj.inline_.extent.cx = cx;
            obj.inline_.graphic.graphicData.pic.spPr.cx = cx;
        end

        % ---- type (@property read-only, shape.py 72-90) ----
        function value = get.type(obj)
            % Python:
            %   graphicData = self._inline.graphic.graphicData
            %   uri = graphicData.uri
            %   if uri == nsmap["pic"]:
            %       blip = graphicData.pic.blipFill.blip
            %       if blip.link is not None: return WD_INLINE_SHAPE.LINKED_PICTURE
            %       return WD_INLINE_SHAPE.PICTURE
            %   if uri == nsmap["c"]:   return WD_INLINE_SHAPE.CHART
            %   if uri == nsmap["dgm"]: return WD_INLINE_SHAPE.SMART_ART
            %   return WD_INLINE_SHAPE.NOT_IMPLEMENTED
            % H10: uri->enum dispatch by fixed nsmap URIs. H3: blip.link is None
            % check -> isequal(link, []).
            m = mat2doc.oxml.nsmap();
            graphicData = obj.inline_.graphic.graphicData;
            uri = graphicData.uri;
            if uri == m.pic
                blip = graphicData.pic.blipFill.blip;
                if ~isequal(blip.link, [])   % Python: if blip.link is not None
                    value = mat2doc.enum.shape.WD_INLINE_SHAPE.LINKED_PICTURE;
                    return
                end
                value = mat2doc.enum.shape.WD_INLINE_SHAPE.PICTURE;
                return
            end
            if uri == m.c
                value = mat2doc.enum.shape.WD_INLINE_SHAPE.CHART;
                return
            end
            if uri == m.dgm
                value = mat2doc.enum.shape.WD_INLINE_SHAPE.SMART_ART;
                return
            end
            value = mat2doc.enum.shape.WD_INLINE_SHAPE.NOT_IMPLEMENTED;
        end
    end
end
