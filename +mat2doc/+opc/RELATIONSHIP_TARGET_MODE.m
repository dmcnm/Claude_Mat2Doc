classdef RELATIONSHIP_TARGET_MODE
% RELATIONSHIP_TARGET_MODE Open XML relationship target modes.
%
%   Access as mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL / .EXTERNAL. Hard
%   dependency of CT_Relationship.new (its target_mode default is INTERNAL) and
%   of CT_Relationships.add_rel (chooses EXTERNAL vs INTERNAL from is_external).
%
%   Example:
%       disp(mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL)   % "Internal"
%       disp(mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL)   % "External"
%
%   Ported from python-docx v1.2.0: src/docx/opc/constants.py::
%   RELATIONSHIP_TARGET_MODE (lines 171-175)

    properties (Constant)
        EXTERNAL = "External"
        INTERNAL = "Internal"
    end
end
