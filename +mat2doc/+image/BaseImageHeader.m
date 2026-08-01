classdef BaseImageHeader < handle
% BASEIMAGEHEADER Base class for image header subclasses (Png, Gif, Bmp, ...).
%
%   Holds the four characteristics every image header exposes -- px_width,
%   px_height, horz_dpi, vert_dpi -- as read-only Dependent properties over
%   protected storage. content_type and default_ext are abstract in Python
%   (@property raising NotImplementedError); here they are METHODS the concrete
%   subclasses override (a MATLAB subclass cannot redefine an inherited
%   property's getter, so the overridable members are methods, while the four
%   stored characteristics remain inherited Dependent properties).
%
%   Ported from python-docx v1.2.0: src/docx/image/image.py::BaseImageHeader
%   (lines 185-234)

    properties (Access = protected)
        px_width_
        px_height_
        horz_dpi_
        vert_dpi_
    end

    properties (Dependent)
        px_width      % horizontal pixel dimension
        px_height     % vertical pixel dimension
        horz_dpi      % integer dots-per-inch, width
        vert_dpi      % integer dots-per-inch, height
    end

    methods
        function obj = BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi)
            % Ported from BaseImageHeader.__init__ (image.py 188-192).
            obj.px_width_ = px_width;
            obj.px_height_ = px_height;
            obj.horz_dpi_ = horz_dpi;
            obj.vert_dpi_ = vert_dpi;
        end

        function v = get.px_width(obj)
            % px_width @property (image.py 210-213).
            v = obj.px_width_;
        end

        function v = get.px_height(obj)
            % px_height @property (image.py 215-218).
            v = obj.px_height_;
        end

        function v = get.horz_dpi(obj)
            % horz_dpi @property (image.py 220-226).
            v = obj.horz_dpi_;
        end

        function v = get.vert_dpi(obj)
            % vert_dpi @property (image.py 228-234).
            v = obj.vert_dpi_;
        end

        function ct = content_type(obj) %#ok<STOUT,MANU>
            % content_type @property (image.py 194-198): abstract; concrete
            %   subclasses (Png/Gif/Bmp) override this method.
            error("mat2doc:NotImplementedError", "%s", ...
                "content_type property must be implemented by all subclasses of BaseImageHeader");
        end

        function e = default_ext(obj) %#ok<STOUT,MANU>
            % default_ext @property (image.py 200-208): abstract; concrete
            %   subclasses (Png/Gif/Bmp) override this method.
            error("mat2doc:NotImplementedError", "%s", ...
                "default_ext property must be implemented by all subclasses of BaseImageHeader");
        end
    end
end
