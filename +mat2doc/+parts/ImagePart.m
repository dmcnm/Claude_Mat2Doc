classdef ImagePart < mat2doc.opc.Part
% IMAGEPART An image part (target of a RELATIONSHIP_TYPE.IMAGE relationship).
%
%   A base Part specialization for image binaries. Its blob is the image bytes,
%   returned VERBATIM (it inherits Part.blob unchanged, so image parts round-trip
%   byte-for-byte). Adds the docx image surface: default_cx / default_cy (native
%   EMU dimensions from the image pixel size + dpi), filename, image (lazy
%   Image.from_blob), sha1, and the from_image / load classmethods.
%
%   ARG ORDER (docx, differs from python-pptx): docx ImagePart.__init__ is
%   (partname, content_type, blob, image=None) -- image LAST, and NO package
%   argument (it forwards only partname/content_type/blob to Part, leaving the
%   base package = None). python-pptx's ImagePart took (partname, content_type,
%   package, blob, filename); the docx shape is preserved here.
%
%   LAZY image (image.py 65-69): NOT a @lazyproperty flag pattern -- python-docx
%   uses the manual `if self._image is None` sentinel, ported verbatim as the
%   None-idiom isequal(image_, []) test (an Image handle is never [], so the test
%   is sound, H3). from_image seeds image_ eagerly; load / PartFactory leave it []
%   so the first .image access decodes the blob.
%
%   FACTORY DISPATCH (H10): PartFactory.part_class_selector_ maps reltype IMAGE ->
%   "mat2doc.parts.ImagePart" (the P7-4 flip from the base mat2doc.opc.Part
%   stand-in). `load` is the factory entry point (MATLAB does not dispatch an
%   inherited static method to the subclass -- the inherited-static trap, same as
%   StylesPart/HeaderPart), constructing an ImagePart with the verbatim blob so a
%   LOADED image part answers .sha1 (required by the ImageParts sha1 dedupe).
%
%   default_cx / default_cy (image.py 29-45): native EMU dimensions. NOTE the
%   python-docx source uses horz_dpi for BOTH cx and cy (image.py:43 reads
%   horz_dpi in default_cy, not vert_dpi) -- ported VERBATIM (design.md section 7:
%   no behavior improvement). These members are part of the public API but are NOT
%   on the add_picture path (that uses Image.scaled_dimensions); they transit the
%   lazy `image` decode.
%
%   UNDERSCORE ROTATION (design.md section 2): private `_image` -> image_.
%
%   Example:
%       CT  = mat2doc.opc.CONTENT_TYPE;
%       ip  = mat2doc.parts.ImagePart.load( ...
%           mat2doc.opc.PackURI("/word/media/image1.png"), CT.PNG, blob, []);
%       disp(ip.filename())   % "image.png"          (no source Image)
%       disp(ip.sha1())       % 40-char hex digest
%
%   Ported from python-docx v1.2.0: src/docx/parts/image.py::ImagePart

    properties (Access = private)
        image_                               % _image (an Image, or [] = None)
    end

    methods
        function obj = ImagePart(partname, content_type, blob, image)
            % Ported from ImagePart.__init__ (image.py 23-27): pass-through to
            %   Part(partname, content_type, blob) -- NO package arg (docx forwards
            %   only three args; the base package defaults to None) -- then store
            %   image (default None).
            arguments
                partname
                content_type
                blob
                image = []          % Python image: Image | None = None (H3)
            end
            % Python: super(ImagePart, self).__init__(partname, content_type, blob)
            obj@mat2doc.opc.Part(partname, content_type, blob);
            obj.image_ = image;   % Python: self._image = image
        end

        function cx = default_cx(obj)
            % default_cx @property (image.py 29-36): native width, from the image
            %   pixel width and horizontal dpi. Python:
            %     px_width = self.image.px_width
            %     horz_dpi = self.image.horz_dpi
            %     width_in_inches = px_width / horz_dpi
            %     return Inches(width_in_inches)
            %   px_width / horz_dpi is Python true division; Inches() truncates via
            %   int() (H6). Transits the lazy `image` decode.
            px_width = obj.image().px_width;
            horz_dpi = obj.image().horz_dpi;
            width_in_inches = px_width / horz_dpi;
            cx = mat2doc.shared.Inches(width_in_inches);
        end

        function cy = default_cy(obj)
            % default_cy @property (image.py 38-45): native height, from the image
            %   pixel height and (VERBATIM) HORIZONTAL dpi. Python:
            %     px_height = self.image.px_height
            %     horz_dpi = self.image.horz_dpi        # <- horz_dpi, not vert_dpi
            %     height_in_emu = int(round(914400 * px_height / horz_dpi))
            %     return Emu(height_in_emu)
            %   int(round(...)) is Python round-half-to-even then truncate-toward-
            %   zero (H6): pyRound_ (file-local) then fix(). Emu holds the exact int.
            px_height = obj.image().px_height;
            horz_dpi = obj.image().horz_dpi;      % docx image.py:43 uses horz_dpi
            height_in_emu = fix(pyRound_(914400 * px_height / horz_dpi));
            cy = mat2doc.shared.Emu(height_in_emu);
        end

        function fn = filename(obj)
            % filename @property (image.py 47-57): the source filename, or a
            %   generic "image.<ext>" when the part was built from an unnamed
            %   stream (the default extension follows the detected MIME type via
            %   the partname). Python:
            %     if self._image is not None:
            %         return self._image.filename
            %     return "image.%s" % self.partname.ext
            if ~isequal(obj.image_, [])   % Python: if self._image is not None
                fn = obj.image_.filename;
                return
            end
            % Python: "image.%s" % self.partname.ext
            fn = "image." + obj.partname().ext;
        end

        function img = image(obj)
            % image @property (image.py 65-69): the |Image| for this part's blob,
            %   decoded lazily on first access. Python:
            %     if self._image is None:
            %         self._image = Image.from_blob(self.blob)
            %     return self._image
            %   Manual None-sentinel cache (NOT @lazyproperty), ported verbatim.
            if isequal(obj.image_, [])   % Python: if self._image is None
                obj.image_ = mat2doc.image.Image.from_blob(obj.blob());
            end
            img = obj.image_;
        end

        function v = sha1(obj)
            % sha1 @property (image.py 77-80): the SHA1 hex digest of the blob.
            %   Python: hashlib.sha1(self.blob).hexdigest(). A plain @property
            %   (recomputed each call; docx does NOT cache it, unlike pptx).
            v = mat2doc.opc.sha1_hexdigest(obj.blob());
        end
    end

    methods (Static)
        function obj = from_image(image, partname)
            % from_image @classmethod (image.py 59-63): a new ImagePart built from
            %   `image` and assigned `partname`. Python:
            %     return ImagePart(partname, image.content_type, image.blob, image)
            obj = mat2doc.parts.ImagePart( ...
                partname, image.content_type, image.blob, image);
        end

        function obj = load(partname, content_type, blob, package) %#ok<INUSD>
            % load @classmethod (image.py 71-75): the PartFactory entry point for
            %   loading an image part from a package being opened. Python:
            %     return cls(partname, content_type, blob)
            %   `package` is accepted (PartFactory calls load with it) but NOT
            %   forwarded -- docx image parts carry no package back-reference.
            %   Declared here because MATLAB does not dispatch an inherited static
            %   method to the subclass (the inherited-static trap; StylesPart /
            %   HeaderPart precedent).
            obj = mat2doc.parts.ImagePart(partname, content_type, blob);
        end
    end
end

% ------------------------------------------------------------------------
% File-local helper: Python 3 one-argument round() (half-to-even), used only by
% default_cy's int(round(...)) (H6). Mirrors the package-private pyRound copies
% in +image/private and +oxml/+simpletypes/private (each package keeps its own;
% +parts has no private/ dir and default_cy is the sole consumer here).
% ------------------------------------------------------------------------
function y = pyRound_(x)
r = floor(x);
d = x - r;
if d < 0.5
    y = r;
elseif d > 0.5
    y = r + 1;
else
    if mod(r, 2) == 0
        y = r;
    else
        y = r + 1;
    end
end
y = y + 0;   % normalize IEEE -0.0 to +0.0
end
