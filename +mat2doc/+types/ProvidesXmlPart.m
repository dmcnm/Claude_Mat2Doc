classdef (Abstract) ProvidesXmlPart < handle
% PROVIDESXMLPART Structural type: an object that provides access to its XmlPart.
%
%   This type is for objects that need access to their part but it either isn't
%   a StoryPart or they don't care, possibly because they just need access to
%   the package or related parts.
%
%   STRUCTURAL / DUCK-TYPED (audit_P2-1_proxy_tier.md): in python-docx this is a
%   typing.Protocol (types.py 25-34) -- a STRUCTURAL type used only for static
%   type checking. It has no runtime identity: an object satisfies it merely by
%   exposing a `part` accessor, WITHOUT inheriting from it. Python-docx never
%   does `isinstance(x, ProvidesXmlPart)`.
%
%   Accordingly this MATLAB class is DOCUMENTATION-ONLY. It declares the
%   `part` contract as an Abstract method so `help`/the class browser show the
%   required surface, but NOTHING in the port inherits from it and NOTHING does
%   `isa(x, "mat2doc.types.ProvidesXmlPart")`. The contract is satisfied
%   structurally: any object with a `part()` method (e.g. mat2doc.opc.Part and
%   its subclasses, or an ElementProxy/Parented whose parent provides one) is a
%   ProvidesXmlPart. This mirrors the Mat2Ppt precedent, where pptx/types.py's
%   Protocols were carried as documentation (referenced as "a ProvidesPart" in
%   member help) with no enforcing class.
%
%   Do NOT subclass this to gain behavior and do NOT gate logic on it; it exists
%   solely to document the duck-typed `part`-provider contract named in the
%   ElementProxy / Parented signatures.
%
%   Ported from python-docx v1.2.0: src/docx/types.py::ProvidesXmlPart

    methods (Abstract)
        % PART The package part containing / reachable from this object -- an
        %   XmlPart (types.py 33-34, @property). Ported as a zero-argument
        %   method (house convention). STRUCTURAL: implemented by any part
        %   provider without deriving from this class.
        p = part(obj)
    end
end
