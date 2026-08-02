classdef Drawing < mat2doc.shared.Parented
% DRAWING Container for a DrawingML object (a `<w:drawing>` element).
%
%   A drawing can contain a picture, but it can also contain a chart, SmartArt,
%   or a drawing canvas. Picture-related members (`.image`) raise when the
%   drawing does not contain a picture; use `.has_picture` to qualify a Drawing
%   before using them.
%
%   This is the LAST catalog module ported (C2b -- the one docx module with no
%   prior WP). It is a pure API/proxy tier over the already-registered CT_Drawing
%   (oxml/drawing.py, P7-3): it adds NO oxml logic, NO registry rows and NO
%   serialization code. Equivalence is BEHAVIORAL.
%
%   TIER (drawing/__init__.py 15 `class Drawing(Parented)`): Parented < handle
%   (the parent-ONLY tier, P2-1) -- it holds NO element and does NOT define
%   eq/ne, so a Drawing is compared by MATLAB's DEFAULT handle identity
%   (instance identity) == Python's default object identity, NOT by wrapped-
%   element identity (H5). Because Parented holds no element_, this class
%   declares its OWN private handles (drawing_ / element_).
%
%   ATTRIBUTES (drawing/__init__.py 18-21): Python `super().__init__(parent);
%   self._parent = parent; self._drawing = self._element = drawing`. Parented
%   already stored parent_; the redundant re-assignment is a no-op here (parent_
%   is inherited, protected). drawing_ (Python _drawing, the working handle) and
%   element_ (Python _element, set but never READ inside drawing/__init__.py --
%   ported for fidelity) both hold the same CT_Drawing.
%
%   HAS_PICTURE (drawing/__init__.py 23-44): true when the drawing contains an
%   embedded (non-linked) picture, inline OR floating. Python
%   `bool(self._drawing.xpath(xpath_expr))` -> ~isempty of the xpath result (H4:
%   an empty match list is falsy). The union `|` + `.` child paths are the WP5
%   xpath subset (design.md section 3). False for a linked picture, chart,
%   SmartArt, or canvas.
%
%   IMAGE (drawing/__init__.py 46-59): an Image proxy for the picture's image.
%   Raises mat2doc:ValueError (VERBATIM Python "drawing does not contain a
%   picture") when the drawing is not a picture. Python:
%     picture_rIds = self._drawing.xpath(".//pic:blipFill/a:blip/@r:embed")
%     if not picture_rIds: raise ValueError("drawing does not contain a picture")
%     rId = picture_rIds[0]
%     doc_part = self.part
%     image_part = doc_part.related_parts[rId]
%     return image_part.image
%   H4: `if not picture_rIds` -> isempty (empty match list is falsy). H1: the
%   `[0]` first element -> (1) (xpath `/@attr` positions are DATA, not a shifted
%   index). related_parts is the {rId -> target Part} map whose values are 1x1
%   cells (P1-5 currency); [rId] -> unwrap the cell (H5: the live image part
%   handle). Ported as a Dependent read-only property to preserve the Python
%   @property shape; MATLAB auto-display catches the ValueError and omits the row
%   when the drawing is not a picture.
%
%   Example:
%       % Reached via Run.iter_inner_content when a run holds a <w:drawing>:
%       items = run.iter_inner_content();
%       d = items{1};                 % a mat2doc.drawing.Drawing (say)
%       if d.has_picture
%           img = d.image;            % a mat2doc.image.Image
%       end
%
%   Ported from python-docx v1.2.0: src/docx/drawing/__init__.py::Drawing

    properties (Access = private)
        drawing_    % _drawing (drawing/__init__.py 21): the working <w:drawing> (a CT_Drawing)
        element_    % _element (drawing/__init__.py 21): same handle; set but never read
    end

    properties (Dependent)
        has_picture   % bool -- true when the drawing contains an embedded picture
        image         % mat2doc.image.Image -- the picture's image (raises if not a picture)
    end

    methods
        function obj = Drawing(drawing, parent)
            % DRAWING Wrap a `<w:drawing>` element (drawing/__init__.py 18-21).
            %
            %   Inputs:  drawing - a mat2doc.oxml.drawing.CT_Drawing (the
            %                      `w:drawing` element).
            %            parent  - the parent proxy (a ProvidesStoryPart)
            %                      providing `part`.
            %   Outputs: obj     - a scalar Drawing handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/drawing/__init__.py::Drawing.__init__
            obj@mat2doc.shared.Parented(parent);   % Python: super().__init__(parent)
            % Python: self._parent = parent (parent_ already set by Parented; no-op)
            % Python: self._drawing = self._element = drawing (one element, two names)
            obj.drawing_ = drawing;
            obj.element_ = drawing;
        end

        % ============================ has_picture ============================
        function value = get.has_picture(obj)
            % HAS_PICTURE true when this drawing contains an embedded picture,
            %   inline OR floating (drawing/__init__.py 23-44). Python:
            %     xpath_expr = ("./wp:inline/a:graphic/a:graphicData/pic:pic"
            %                   " | ./wp:anchor/a:graphic/a:graphicData/pic:pic")
            %     return bool(self._drawing.xpath(xpath_expr))
            xpath_expr = "./wp:inline/a:graphic/a:graphicData/pic:pic" + ...
                " | ./wp:anchor/a:graphic/a:graphicData/pic:pic";
            value = ~isempty(obj.drawing_.xpath(xpath_expr));   % Python: bool(...xpath(...))
        end

        % ============================ image ============================
        function value = get.image(obj)
            % IMAGE An Image proxy for the picture in this (picture) drawing
            %   (drawing/__init__.py 46-59). Raises mat2doc:ValueError when the
            %   drawing does not contain a picture.
            %
            %   Ported from python-docx v1.2.0: src/docx/drawing/__init__.py::Drawing.image
            picture_rIds = obj.drawing_.xpath(".//pic:blipFill/a:blip/@r:embed");
            if isempty(picture_rIds)   % Python: if not picture_rIds (empty list falsy, H4)
                error("mat2doc:ValueError", "%s", "drawing does not contain a picture");
            end
            rId = picture_rIds(1);              % Python: rId = picture_rIds[0]
            doc_part = obj.part();              % Python: doc_part = self.part
            rp = doc_part.related_parts();      % Python: doc_part.related_parts
            cellval = rp(rId);                  % Python: related_parts[rId]
            image_part = cellval{1};            % unwrap the P1-5 1x1-cell currency (H5)
            value = image_part.image;           % Python: return image_part.image
        end
    end
end
