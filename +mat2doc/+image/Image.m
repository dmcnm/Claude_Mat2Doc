classdef Image < handle
% IMAGE Graphical image stream characterized as to content type and size.
%
%   A JPEG, PNG, GIF, BMP, or TIFF image stream with the properties and methods
%   required by ImagePart. It decodes an image blob's header to expose
%   content_type, px_width/px_height, horz_dpi/vert_dpi, width/height, sha1 and
%   ext -- the characterization docx needs before adding an image to a document.
%
%   HANDLE (design.md section 2): the derived properties ext and sha1 are
%   @lazyproperty (cache-on-first-access), so this is a handle class with the
%   private-cache + logical-computed-flag idiom.
%
%   UNDERSCORE ROTATION (design.md section 2): _blob -> blob_, _filename ->
%   filename_, _image_header -> image_header_, _from_stream -> from_stream_.
%
%   NONE IDIOM (H3): python-docx uses None for an absent filename / scaled
%   dimension; Mat2Doc has no shared isNone helper, so the None-guard is the
%   inline isequal(x, []) test (ratified none-idiom decision 2026-07-26).
%
%   Ported from python-docx v1.2.0: src/docx/image/image.py::Image (lines 18-165)

    properties (Access = private)
        blob_                                % _blob (image binary bytestream)
        filename_                            % _filename (str or None/[])
        image_header_                        % _image_header (a BaseImageHeader)
        ext_cache_
        ext_isComputed_ (1,1) logical = false
        sha1_cache_
        sha1_isComputed_ (1,1) logical = false
    end

    properties (Dependent)
        blob          % @property: binary image bytestream
        content_type  % @property: MIME content type from the header
        ext           % @lazyproperty: file extension from the load filename
        filename      % @property: source filename, or None
        px_width      % @property: horizontal pixel dimension
        px_height     % @property: vertical pixel dimension
        horz_dpi      % @property: integer dots-per-inch, width
        vert_dpi      % @property: integer dots-per-inch, height
        width         % @property: native width as a Length (Inches)
        height        % @property: native height as a Length (Inches)
        sha1          % @lazyproperty: SHA1 hex digest of the blob
    end

    methods
        function obj = Image(blob, filename, image_header)
            % Ported from Image.__init__ (image.py 22-26).
            obj.blob_ = blob;
            obj.filename_ = filename;
            obj.image_header_ = image_header;
        end

        function v = get.blob(obj)
            % blob @property (image.py 52-55).
            v = obj.blob_;
        end

        function v = get.content_type(obj)
            % content_type @property (image.py 57-60): from the image header.
            v = obj.image_header_.content_type;
        end

        function v = get.ext(obj)
            % ext @lazyproperty (image.py 62-70):
            %   os.path.splitext(self._filename)[1][1:] -- the extension after
            %   the last dot in the basename, WITHOUT the leading period and
            %   WITHOUT lowercasing. MATLAB's fileparts does NOT match CPython
            %   os.path.splitext on dotfiles (it treats ".bashrc" as an
            %   extension), so the splitext extension rule is replicated in the
            %   package-private helper splitext_ext (Gate-2 F-1 fix). filename_
            %   is always set to a real string by from_stream_
            %   ("image.<default_ext>" when the loader supplies none), so the
            %   None-filename TypeError path Python would take is unreachable on
            %   a constructed Image.
            if ~obj.ext_isComputed_
                obj.ext_cache_ = splitext_ext(obj.filename_);  % returns a string
                obj.ext_isComputed_ = true;
            end
            v = obj.ext_cache_;
        end

        function v = get.filename(obj)
            % filename @property (image.py 72-76).
            v = obj.filename_;
        end

        function v = get.px_width(obj)
            % px_width @property (image.py 78-81): from the image header.
            v = obj.image_header_.px_width;
        end

        function v = get.px_height(obj)
            % px_height @property (image.py 83-86): from the image header.
            v = obj.image_header_.px_height;
        end

        function v = get.horz_dpi(obj)
            % horz_dpi @property (image.py 88-94): from the image header.
            v = obj.image_header_.horz_dpi;
        end

        function v = get.vert_dpi(obj)
            % vert_dpi @property (image.py 96-102): from the image header.
            v = obj.image_header_.vert_dpi;
        end

        function v = get.width(obj)
            % width @property (image.py 104-108): Inches(px_width / horz_dpi).
            %   Python true division; Inches() truncates via int() (H6).
            v = mat2doc.shared.Inches(obj.px_width / obj.horz_dpi);
        end

        function v = get.height(obj)
            % height @property (image.py 110-114): Inches(px_height / vert_dpi).
            v = mat2doc.shared.Inches(obj.px_height / obj.vert_dpi);
        end

        function v = get.sha1(obj)
            % sha1 @lazyproperty (image.py 148-151).
            if ~obj.sha1_isComputed_
                obj.sha1_cache_ = mat2doc.opc.sha1_hexdigest(obj.blob_);
                obj.sha1_isComputed_ = true;
            end
            v = obj.sha1_cache_;
        end

        function [cx, cy] = scaled_dimensions(obj, width, height)
            % scaled_dimensions (image.py 116-146): (cx, cy) scaled Length pair.
            %   Both None -> native (width, height); one None -> computed from
            %   the other preserving aspect ratio; both set -> (width, height).
            %   Returns two Emu Length values. None-guard is the inline
            %   isequal(x, []) idiom (H3).
            arguments
                obj
                width = []     % None (H3)
                height = []    % None (H3)
            end
            if isequal(width, []) && isequal(height, [])
                cx = obj.width;
                cy = obj.height;
                return
            end

            if isequal(width, [])
                % assert height is not None
                scaling_factor = double(height) / double(obj.height);
                width = pyRound(double(obj.width) * scaling_factor);  % round(..) H6
            end

            if isequal(height, [])
                scaling_factor = double(width) / double(obj.width);
                height = pyRound(double(obj.height) * scaling_factor);
            end

            cx = mat2doc.shared.Emu(width);
            cy = mat2doc.shared.Emu(height);
        end
    end

    methods (Static)
        function obj = from_blob(blob)
            % from_blob (image.py 28-33): parse an Image from the image binary.
            stream = mat2doc.image.BytesIO(blob);
            obj = mat2doc.image.Image.from_stream_(stream, blob);
        end

        function obj = from_file(image_descriptor)
            % from_file (image.py 35-50): load from a path (str/char) or a
            %   file-like object (a BytesIO).
            if isstring(image_descriptor) || ischar(image_descriptor)
                path = image_descriptor;
                fid = fopen(path, "r");                 % MATLAB 'r' is binary
                if fid < 0
                    error("mat2doc:FileNotFoundError", ...
                        "cannot open image file '%s'", string(path));
                end
                c = onCleanup(@() fclose(fid));
                blob = fread(fid, inf, "*uint8")';
                clear c                                  % fclose now
                stream = mat2doc.image.BytesIO(blob);
                [~, nm, xt] = fileparts(char(path));
                filename = string(nm) + string(xt);      % os.path.basename
            else
                stream = image_descriptor;               % file-like object
                stream.seek(0);
                blob = stream.read();                    % read all remaining
                filename = [];                           % None
            end
            obj = mat2doc.image.Image.from_stream_(stream, blob, filename);
        end
    end

    methods (Static, Access = private)
        function obj = from_stream_(stream, blob, filename)
            % _from_stream (image.py 153-165): dispatch a header for `stream`,
            %   default the filename to "image.<default_ext>" when none given.
            arguments
                stream
                blob
                filename = []    % None (H3)
            end
            image_header = mat2doc.image.ImageHeaderFactory_(stream);
            if isequal(filename, [])
                filename = "image." + image_header.default_ext;   % "image.%s"
            end
            obj = mat2doc.image.Image(blob, filename, image_header);
        end
    end
end
