classdef StoryChild < handle
% STORYCHILD A document element within a story part.
%
%   Story parts include DocumentPart and Header/FooterPart and can contain
%   block items (paragraphs and tables). Items from the block-item subtree
%   occasionally require an ancestor object to provide access to part-level or
%   package-level items like styles or images, or to add or drop a
%   relationship. Provides the `parent_` attribute to subclasses.
%
%   DOCX STRUCTURE (audit_P2-1_proxy_tier.md): StoryChild is structurally
%   identical to Parented -- it holds only a parent and delegates `part` to it
%   -- but is a DISTINCT class in python-docx (shared.py 336-353), because its
%   parent is typed as a ProvidesStoryPart (its `part` is a StoryPart) rather
%   than a bare ProvidesXmlPart. The two are kept as separate MATLAB classes to
%   preserve the docx class hierarchy exactly; block-item subclasses (Paragraph,
%   Table, ...) derive from StoryChild, not Parented. Like Parented it does NOT
%   override eq/ne (default handle identity == Python default object identity).
%
%   PROPERTY-AS-METHOD (design.md section 2): Python's `part` @property is
%   ported as a zero-argument method (house convention).
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_parent` attribute
%   rotates the leading underscore -> parent_.
%
%   Example:
%       % An ancestor is any object exposing a `part` accessor whose part is a
%       % StoryPart (a ProvidesStoryPart in Python).
%       obj = mat2doc.shared.StoryChild(some_story_ancestor);
%       p   = obj.part();     % delegated up: parent_.part() (a StoryPart)
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::StoryChild

    properties (Access = protected)
        parent_         % _parent: the object providing `part` (a ProvidesStoryPart)
    end

    methods
        function obj = StoryChild(parent)
            % STORYCHILD Store the parent object (shared.py 347-348):
            %   self._parent = parent.
            %
            %   Inputs:  parent - an object providing a `part` accessor whose
            %                     part is a StoryPart (a ProvidesStoryPart).
            %   Outputs: obj    - a scalar StoryChild handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::StoryChild.__init__
            obj.parent_ = parent;
        end

        function p = part(obj)
            % PART The package part containing this object (shared.py 350-353):
            %   delegates to the parent -- `self._parent.part` (a StoryPart).
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::StoryChild.part
            p = obj.parent_.part();
        end
    end
end
