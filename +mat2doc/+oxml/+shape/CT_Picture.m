classdef CT_Picture < mat2doc.oxml.BaseOxmlElement
% CT_PICTURE Custom element class for the <pic:pic> element.
%
%   `<pic:pic>` element, a DrawingML picture. Three OneAndOnlyOne required
%   children -- nvPicPr (pic:nvPicPr), blipFill (pic:blipFill), spPr (pic:spPr)
%   -- plus the `new(pic_id, filename, rId, cx, cy)` builder classmethod and its
%   `_pic_xml` template.
%
%   DOCX vs pptx: python-docx's CT_Picture is the INLINE-picture pic:pic (pic:*
%   prefix children, a minimal `new` builder). python-pptx's same-named class
%   (oxml/shapes/picture.py) is the slide-picture p:pic (BaseShapeElement, p:*
%   children, new_pic/new_ph_pic/new_video_pic, srcRect/crop/ln surface). Ported
%   to match DOCX (pic: children, the single `new` builder). Registered for
%   <pic:pic> in docx (oxml/__init__.py:56).
%
%   ============================ BYTE-CRITICAL BUILDER ============================
%   new(pic_id, filename, rId, cx, cy) (shape.py:146-155) parse_xml's the EXACT
%   python-docx _pic_xml template then stamps the five slots:
%     pic.nvPicPr.cNvPr.id   = pic_id       (@id  on pic:cNvPr, template "666")
%     pic.nvPicPr.cNvPr.name = filename     (@name on pic:cNvPr, template "unnamed")
%     pic.blipFill.blip.embed = rId         (@r:embed on a:blip, template absent)
%     pic.spPr.cx = cx                      (a:xfrm/a:ext/@cx, template "914400")
%     pic.spPr.cy = cy                      (a:xfrm/a:ext/@cy, template "914400")
%   The template nsdecls order is ("pic","a","r") (H8: fresh nsdecls emitted in
%   template order). This feeds CT_Inline.new_pic_inline (P7-4 add_picture). The
%   built <pic:pic> bytes must be byte-identical to python-docx.
%
%   UNDERSCORE ROTATION (design.md section 2): private _pic_xml -> pic_xml_.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shape.py::CT_Picture
%   (lines 135-179; registered for <pic:pic>, oxml/__init__.py:56)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        NVPICPR_TAG = "pic:nvPicPr"    % OneAndOnlyOne @ shape.py:138
        BLIPFILL_TAG = "pic:blipFill"  % OneAndOnlyOne @ shape.py:141
        SPPR_TAG = "pic:spPr"          % OneAndOnlyOne @ shape.py:144
    end

    properties (Dependent)  % generated descriptor properties
        nvPicPr    % OneAndOnlyOne('pic:nvPicPr') -- required child (CT_PictureNonVisual)
        blipFill   % OneAndOnlyOne('pic:blipFill') -- required child (CT_BlipFillProperties)
        spPr       % OneAndOnlyOne('pic:spPr') -- required child (CT_ShapeProperties)
    end

    methods
        function obj = CT_Picture(varargin)
            % CT_PICTURE Construct a loose <pic:pic> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- nvPicPr (OneAndOnlyOne) ----
        function child = get.nvPicPr(obj)
            child = obj.getRequiredChild(obj.NVPICPR_TAG);
        end

        % ---- blipFill (OneAndOnlyOne) ----
        function child = get.blipFill(obj)
            child = obj.getRequiredChild(obj.BLIPFILL_TAG);
        end

        % ---- spPr (OneAndOnlyOne) ----
        function child = get.spPr(obj)
            child = obj.getRequiredChild(obj.SPPR_TAG);
        end
    end

    methods (Static)
        % ---- new (classmethod, shape.py:146-155) ----
        function pic = new(pic_id, filename, rId, cx, cy)
            % NEW A new minimum viable <pic:pic> (picture) element.
            %   Python:
            %     pic = parse_xml(cls._pic_xml())
            %     pic.nvPicPr.cNvPr.id = pic_id
            %     pic.nvPicPr.cNvPr.name = filename
            %     pic.blipFill.blip.embed = rId
            %     pic.spPr.cx = cx
            %     pic.spPr.cy = cy
            %     return pic
            pic = mat2doc.oxml.parse_xml(mat2doc.oxml.shape.CT_Picture.pic_xml_());
            pic.nvPicPr.cNvPr.id = pic_id;
            pic.nvPicPr.cNvPr.name = filename;
            pic.blipFill.blip.embed = rId;
            pic.spPr.cx = cx;
            pic.spPr.cy = cy;
        end
    end

    methods (Static, Access = private)
        % ---- _pic_xml -> pic_xml_ (classmethod, shape.py:157-179) ----
        function xml = pic_xml_()
            % PIC_XML_ The exact python-docx <pic:pic> template (shape.py:158-179).
            %   nsdecls order ("pic","a","r") reproduced verbatim (H8).
            nl = newline;
            xml = "<pic:pic " + mat2doc.oxml.nsdecls("pic", "a", "r") + ">" + nl + ...
                "  <pic:nvPicPr>" + nl + ...
                "    <pic:cNvPr id=""666"" name=""unnamed""/>" + nl + ...
                "    <pic:cNvPicPr/>" + nl + ...
                "  </pic:nvPicPr>" + nl + ...
                "  <pic:blipFill>" + nl + ...
                "    <a:blip/>" + nl + ...
                "    <a:stretch>" + nl + ...
                "      <a:fillRect/>" + nl + ...
                "    </a:stretch>" + nl + ...
                "  </pic:blipFill>" + nl + ...
                "  <pic:spPr>" + nl + ...
                "    <a:xfrm>" + nl + ...
                "      <a:off x=""0"" y=""0""/>" + nl + ...
                "      <a:ext cx=""914400"" cy=""914400""/>" + nl + ...
                "    </a:xfrm>" + nl + ...
                "    <a:prstGeom prst=""rect""/>" + nl + ...
                "  </pic:spPr>" + nl + ...
                "</pic:pic>";
        end
    end
end
