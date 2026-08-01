classdef PNG_CHUNK_TYPE
% PNG_CHUNK_TYPE PNG chunk type names.
%
%   Access as mat2doc.image.PNG_CHUNK_TYPE.IHDR etc. String values are the
%   4-character ASCII chunk-type tags, byte-identical to python-docx
%   constants.py. Chunk-type comparison is case-SENSITIVE (H15): "pHYs" is not
%   "PHYS"; the values here preserve the exact mixed case.
%
%   Example:
%       disp(mat2doc.image.PNG_CHUNK_TYPE.pHYs)   % "pHYs"
%
%   Ported from python-docx v1.2.0: src/docx/image/constants.py::PNG_CHUNK_TYPE
%   (lines 110-115)

    properties (Constant)
        IHDR = "IHDR"
        pHYs = "pHYs"
        IEND = "IEND"
    end
end
