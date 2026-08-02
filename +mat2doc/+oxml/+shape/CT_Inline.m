classdef CT_Inline < mat2doc.oxml.BaseOxmlElement
% CT_INLINE Custom element class for the <wp:inline> element.
%
%   `<wp:inline>` element, container for an inline shape. Three OneAndOnlyOne
%   required children -- extent (wp:extent), docPr (wp:docPr), graphic
%   (a:graphic) -- plus the two builder classmethods `new` / `new_pic_inline`
%   and the `_inline_xml` template.
%
%   NOVEL (docx wp:-namespace inline-picture structure): python-pptx has no
%   CT_Inline. Registered for <wp:inline> in docx (oxml/__init__.py:62).
%
%   ============================ BYTE-CRITICAL BUILDER ============================
%   new_pic_inline(shape_id, rId, filename, cx, cy) (shape.py:92-103) builds the
%   COMPLETE wp:inline / a:graphic / pic:pic tree for an inline picture. It is the
%   XML builder P7-4's add_picture calls; the M2-style picture byte scenario
%   depends on the bytes being byte-identical to python-docx.
%
%   new(cx, cy, shape_id, pic) (shape.py:79-90) parse_xml's the EXACT
%   _inline_xml template then stamps the slots:
%     inline.extent.cx = cx / inline.extent.cy = cy   (wp:extent, template 914400)
%     inline.docPr.id   = shape_id                     (wp:docPr @id, template 666)
%     inline.docPr.name = "Picture %d" % shape_id      (wp:docPr @name)
%     inline.graphic.graphicData.uri = "...drawingml/2006/picture"
%     inline.graphic.graphicData._insert_pic(pic)      (append the pic:pic subtree)
%   The template nsdecls order is ("wp","a","pic","r") (H8: fresh nsdecls emitted
%   in template order). H11: the wp:inline child sequence is
%   extent -> docPr -> cNvGraphicFramePr -> graphic (fixed by the template), and
%   the pic:pic is appended into a:graphicData (empty in the template) via the
%   generated insert_pic_ (Python _insert_pic, successors=() -> append).
%
%   UNDERSCORE ROTATION (design.md section 2): private _inline_xml -> inline_xml_.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Inline
%   (lines 68-118; registered for <wp:inline>, oxml/__init__.py:62)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        EXTENT_TAG = "wp:extent"    % OneAndOnlyOne @ shape.py:71
        DOCPR_TAG = "wp:docPr"      % OneAndOnlyOne @ shape.py:72
        GRAPHIC_TAG = "a:graphic"   % OneAndOnlyOne @ shape.py:75
        % The graphicData uri set by new() (shape.py:88).
        PICTURE_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    end

    properties (Dependent)  % generated descriptor properties
        extent   % OneAndOnlyOne('wp:extent') -- required child (CT_PositiveSize2D)
        docPr    % OneAndOnlyOne('wp:docPr') -- required child (CT_NonVisualDrawingProps)
        graphic  % OneAndOnlyOne('a:graphic') -- required child (CT_GraphicalObject)
    end

    methods
        function obj = CT_Inline(varargin)
            % CT_INLINE Construct a loose <wp:inline> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- extent (OneAndOnlyOne) ----
        function child = get.extent(obj)
            child = obj.getRequiredChild(obj.EXTENT_TAG);
        end

        % ---- docPr (OneAndOnlyOne) ----
        function child = get.docPr(obj)
            child = obj.getRequiredChild(obj.DOCPR_TAG);
        end

        % ---- graphic (OneAndOnlyOne) ----
        function child = get.graphic(obj)
            child = obj.getRequiredChild(obj.GRAPHIC_TAG);
        end
    end

    methods (Static)
        % ---- new (classmethod, shape.py:79-90) ----
        function inline = new(cx, cy, shape_id, pic)
            % NEW Return a new <wp:inline> element populated with the values passed.
            %   Python:
            %     inline = cast(CT_Inline, parse_xml(cls._inline_xml()))
            %     inline.extent.cx = cx
            %     inline.extent.cy = cy
            %     inline.docPr.id = shape_id
            %     inline.docPr.name = "Picture %d" % shape_id
            %     inline.graphic.graphicData.uri = "...drawingml/2006/picture"
            %     inline.graphic.graphicData._insert_pic(pic)
            %     return inline
            inline = mat2doc.oxml.parse_xml(mat2doc.oxml.shape.CT_Inline.inline_xml_());
            inline.extent.cx = cx;
            inline.extent.cy = cy;
            inline.docPr.id = shape_id;
            % Python: "Picture %d" % shape_id  (H14: %d via pyStr int)
            inline.docPr.name = "Picture " + mat2doc.shared.pyStr(shape_id, "int");
            inline.graphic.graphicData.uri = mat2doc.oxml.shape.CT_Inline.PICTURE_URI;
            inline.graphic.graphicData.insert_pic_(pic);
        end

        % ---- new_pic_inline (classmethod, shape.py:92-103) ----
        function inline = new_pic_inline(shape_id, rId, filename, cx, cy)
            % NEW_PIC_INLINE Create <wp:inline> containing a <pic:pic> element.
            %   The contents of the pic:pic element is taken from the argument
            %   values. Python:
            %     pic_id = 0  # Word doesn't seem to use this, but does not omit it
            %     pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
            %     inline = cls.new(cx, cy, shape_id, pic)
            %     return inline
            pic_id = 0;   % Word doesn't seem to use this, but does not omit it
            pic = mat2doc.oxml.shape.CT_Picture.new(pic_id, filename, rId, cx, cy);
            inline = mat2doc.oxml.shape.CT_Inline.new(cx, cy, shape_id, pic);
        end
    end

    methods (Static, Access = private)
        % ---- _inline_xml -> inline_xml_ (classmethod, shape.py:105-118) ----
        function xml = inline_xml_()
            % INLINE_XML_ The exact python-docx <wp:inline> template
            %   (shape.py:107-118). nsdecls order ("wp","a","pic","r") reproduced
            %   verbatim (H8).
            nl = newline;
            xml = "<wp:inline " + mat2doc.oxml.nsdecls("wp", "a", "pic", "r") + ">" + nl + ...
                "  <wp:extent cx=""914400"" cy=""914400""/>" + nl + ...
                "  <wp:docPr id=""666"" name=""unnamed""/>" + nl + ...
                "  <wp:cNvGraphicFramePr>" + nl + ...
                "    <a:graphicFrameLocks noChangeAspect=""1""/>" + nl + ...
                "  </wp:cNvGraphicFramePr>" + nl + ...
                "  <a:graphic>" + nl + ...
                "    <a:graphicData uri=""URI not set""/>" + nl + ...
                "  </a:graphic>" + nl + ...
                "</wp:inline>";
        end
    end
end
