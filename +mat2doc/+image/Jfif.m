classdef Jfif < mat2doc.image.BaseImageHeader
% JFIF Image header parser for JFIF (JPEG) images. STUB (not yet ported).
%
%   The concrete JFIF BaseImageHeader subclass and the JPEG marker machinery are
%   owned by the P7-2 jpeg WP (which depends on the P7-2 tiff WP; see boundary
%   audit C1). This stub exists so ImageHeaderFactory_'s signature table (which
%   resolves @mat2doc.image.Jfif.from_stream at call time) dispatches to a named
%   target; invoking it raises mat2doc:notYetPorted.
%
%   Target symbol: python-docx v1.2.0 src/docx/image/jpeg.py::Jfif (owning WP: P7-2 jpeg)

    methods (Static)
        function obj = from_stream(stream) %#ok<INUSD,STOUT>
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.image.Jfif.from_stream is not yet ported (owning WP: P7-2 jpeg).");
        end
    end
end
