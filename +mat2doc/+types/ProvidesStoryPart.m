classdef (Abstract) ProvidesStoryPart < handle
% PROVIDESSTORYPART Structural type: an object that provides access to a StoryPart.
%
%   This type is for objects that have a story part like document or header as
%   their root part.
%
%   STRUCTURAL / DUCK-TYPED (audit_P2-1_proxy_tier.md): in python-docx this is a
%   typing.Protocol (types.py 14-22) -- a STRUCTURAL type used only for static
%   type checking. It has no runtime identity: an object satisfies it merely by
%   exposing a `part` accessor whose value is a StoryPart, WITHOUT inheriting
%   from it. Python-docx never does `isinstance(x, ProvidesStoryPart)`.
%
%   Accordingly this MATLAB class is DOCUMENTATION-ONLY. It declares the `part`
%   contract as an Abstract method for discoverability, but NOTHING in the port
%   inherits from it and NOTHING does `isa(x, "mat2doc.types.ProvidesStoryPart")`.
%   The contract is satisfied structurally: any object whose `part()` returns a
%   StoryPart is a ProvidesStoryPart. It is a NARROWER contract than
%   ProvidesXmlPart (the part is specifically a StoryPart); StoryChild names it
%   as its parent's type.
%
%   Do NOT subclass this to gain behavior and do NOT gate logic on it; it exists
%   solely to document the duck-typed StoryPart-provider contract named in the
%   StoryChild signature.
%
%   Ported from python-docx v1.2.0: src/docx/types.py::ProvidesStoryPart

    methods (Abstract)
        % PART The StoryPart containing / reachable from this object (types.py
        %   21-22, @property). Ported as a zero-argument method (house
        %   convention). STRUCTURAL: implemented by any StoryPart provider
        %   without deriving from this class.
        p = part(obj)
    end
end
