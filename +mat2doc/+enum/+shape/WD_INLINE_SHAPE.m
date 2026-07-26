classdef WD_INLINE_SHAPE
% WD_INLINE_SHAPE Alias of WD_INLINE_SHAPE_TYPE.
%
%   Mirrors the python-docx module-level assignment
%   ``WD_INLINE_SHAPE = WD_INLINE_SHAPE_TYPE`` (shape.py line 19). MATLAB has no
%   class aliasing, so this class re-exports the canonical enumeration's members
%   as Constant properties. The members ARE mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE
%   instances, so identity (==) and isa behave exactly as if the two names
%   referred to one enumeration. WD_INLINE_SHAPE_TYPE is a plain enum with no XML
%   mapping, so there are no static methods to forward.
%
%   Example:
%       mat2doc.enum.shape.WD_INLINE_SHAPE.PICTURE == ...
%           mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE   % true
%
%   Ported from python-docx v1.2.0: src/docx/enum/shape.py::WD_INLINE_SHAPE
%   (alias of WD_INLINE_SHAPE_TYPE)

    properties (Constant)
        CHART           = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.CHART
        LINKED_PICTURE  = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.LINKED_PICTURE
        PICTURE         = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.PICTURE
        SMART_ART       = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.SMART_ART
        NOT_IMPLEMENTED = mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE.NOT_IMPLEMENTED
    end
end
