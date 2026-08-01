classdef Gif < mat2doc.image.BaseImageHeader
% GIF Image header parser for GIF images. STUB (not yet ported).
%
%   The concrete GIF BaseImageHeader subclass is owned by P7-1b. This stub exists
%   so ImageHeaderFactory_'s signature table (which resolves
%   @mat2doc.image.Gif.from_stream at call time) dispatches to a named target;
%   invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/gif.py::Gif (owning WP: P7-1b)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Gif.from_stream is not yet ported (owning WP: P7-1b).");
        end
    end
end
