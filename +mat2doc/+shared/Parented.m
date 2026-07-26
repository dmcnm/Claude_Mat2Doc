classdef Parented < handle
% PARENTED Common services for document elements that occur below a part.
%
%   Provides common services for document elements that occur below a part but
%   may occasionally require an ancestor object to provide a service, such as
%   to add or drop a relationship. Provides the `parent_` attribute to
%   subclasses.
%
%   DOCX STRUCTURE (audit_P2-1_proxy_tier.md): unlike python-pptx's
%   ParentedElementProxy, docx's Parented is NOT a subclass of ElementProxy and
%   holds NO element -- only a parent. It therefore exposes only `part`
%   (delegated to the parent) and has NO element-identity `eq`/`ne`. Two
%   Parented instances compare by MATLAB's default handle identity (instance
%   identity), which is exactly Python's default object identity for a class
%   that does not override `__eq__`. So NO eq/ne is defined here.
%
%   PROPERTY-AS-METHOD (design.md section 2): Python's `part` @property is
%   ported as a zero-argument method (house convention).
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_parent` attribute
%   rotates the leading underscore -> parent_.
%
%   Example:
%       % An ancestor is any object exposing a `part` accessor (a
%       % ProvidesXmlPart in Python).
%       obj = mat2doc.shared.Parented(some_ancestor);
%       p   = obj.part();     % delegated up: parent_.part()
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Parented

    properties (Access = protected)
        parent_         % _parent: the object providing `part` (a ProvidesXmlPart)
    end

    methods
        function obj = Parented(parent)
            % PARENTED Store the parent object (shared.py 327-328):
            %   self._parent = parent.
            %
            %   Inputs:  parent - an object providing a `part` accessor
            %                     (a ProvidesXmlPart).
            %   Outputs: obj    - a scalar Parented handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::Parented.__init__
            obj.parent_ = parent;
        end

        function p = part(obj)
            % PART The package part containing this object (shared.py 330-333):
            %   delegates to the parent -- `self._parent.part`.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::Parented.part
            p = obj.parent_.part();
        end
    end
end
