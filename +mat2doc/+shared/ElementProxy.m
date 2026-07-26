classdef ElementProxy < handle
% ELEMENTPROXY Base class for lxml element proxy classes.
%
%   An element proxy class is one whose primary responsibilities are fulfilled
%   by manipulating the attributes and child elements of an XML element. They
%   are the most common type of class in python-docx other than the custom
%   element (oxml) classes. The top-level Document proxy derives from this base
%   (document.py 28: `class Document(ElementProxy)`).
%
%   REFERENCE SEMANTICS (design.md section 2): a handle class. The proxy is a
%   VALUE OBJECT that wraps a shared element tree and holds no mutable local
%   state; two proxy instances that wrap the SAME element are views of one
%   object. This mirrors Python's ElementProxy exactly.
%
%   DOCX-vs-PPTX STRUCTURE (audit_P2-1_proxy_tier.md): unlike python-pptx --
%   where ElementProxy has NO parent and the parent/`part` responsibility lives
%   in the sibling ParentedElementProxy -- python-docx FOLDS the optional parent
%   and the `part` accessor directly into ElementProxy (shared.py 277-316).
%   Parented and StoryChild are SEPARATE, non-derived classes (they hold only a
%   parent, no element). So this class ports docx's ACTUAL shape, not the
%   Mat2Ppt ElementProxy shape; only the MATLAB idiom is borrowed from Mat2Ppt.
%
%   ELEMENT-IDENTITY EQUALITY (H5, shared.py 289-304). `eq`/`ne` are defined so
%   that two proxies compare EQUAL iff they wrap the SAME oxml element --
%   whether or not they are the same proxy instance -- and UNEQUAL against a
%   non-proxy. This is Python's:
%       def __eq__(self, other):
%           if not isinstance(other, ElementProxy): return False
%           return self._element is other._element
%   The element comparison `element_ == element_` delegates to XmlElement's
%   SEALED handle-identity `eq` (design.md section 2): `is` on the underlying
%   node. No content comparison is introduced.
%
%   PROPERTY-AS-METHOD (design.md section 2): Python's `element` and `part`
%   @property members are ported as zero-argument methods, matching the house
%   convention used across the OPC layer (Part.blob, Part.partname, ...).
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_element` /
%   `_parent` attributes rotate the leading underscore -> element_ / parent_.
%
%   NONE SENTINEL (H3): `parent` defaults to [] (the None sentinel). The `part`
%   None-guard tests for that sentinel inline (Mat2Doc uses NO shared isNone
%   helper -- decision_2026-07-26_mat2doc_none_idiom.md). parent_ is only ever
%   [] (a 0x0 double) or a proxy/part handle object, so a double-typed empty
%   check exactly reproduces Python `self._parent is None`.
%
%   NOTE (heterogeneous arrays): unlike XmlElement, ElementProxy is NOT a
%   matlab.mixin.Heterogeneous root and its `eq`/`ne` are NOT Sealed. The proxy
%   layer never forms mixed-class arrays that are compared element-wise; the
%   identity comparison here is always scalar proxy-to-proxy. A downstream
%   collection WP that needs a heterogeneous proxy vector must add the
%   Heterogeneous mixin and seal these methods then (as XmlElement does).
%
%   Example:
%       e   = mat2doc.oxml.XmlElement("w:document");
%       a   = mat2doc.shared.ElementProxy(e);
%       b   = mat2doc.shared.ElementProxy(e);        % different proxy, same element
%       tf  = (a == b);                              % true  (H5 element identity)
%       f   = mat2doc.oxml.XmlElement("w:document");
%       tf2 = (a == mat2doc.shared.ElementProxy(f)); % false (different element)
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy

    properties (Access = protected)
        element_        % _element: the wrapped oxml element (an XmlElement / CT_*)
        parent_         % _parent: the object providing `part`, or [] (None)
    end

    methods
        function obj = ElementProxy(element, parent)
            % ELEMENTPROXY Store the wrapped element and optional parent
            %   (shared.py 285-287): self._element = element; self._parent = parent.
            %
            %   Inputs:  element - the oxml element proxied by this object
            %                      (a mat2doc.oxml.XmlElement or CT_* subclass).
            %            parent  - (optional) an object providing a `part`
            %                      accessor (a ProvidesXmlPart). Default [] (None).
            %   Outputs: obj     - a scalar ElementProxy handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__init__
            arguments
                element
                parent = []     % None sentinel (H3): parent defaults to None
            end
            obj.element_ = element;
            obj.parent_ = parent;
        end

        function e = element(obj)
            % ELEMENT The lxml element proxied by this object (shared.py 306-309).
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.element
            e = obj.element_;
        end

        function p = part(obj)
            % PART The package part containing this object (shared.py 311-316).
            %
            %   Python: if self._parent is None: raise ValueError(...); else
            %   return self._parent.part. The None-guard is ported inline (H3):
            %   parent_ is [] (None) or a handle object with a `part` accessor.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.part
            if isa(obj.parent_, "double") && isempty(obj.parent_)   % _parent is None
                error("mat2doc:ValueError", "%s", ...
                    "part is not accessible from this element");
            end
            p = obj.parent_.part();
        end

        function tf = eq(a, b)
            % EQ True iff a and b wrap the SAME oxml element (shared.py 289-299, H5).
            %
            %   ElementProxy objects are value objects; equality is defined as
            %   referring to the same XML element, whether or not they are the
            %   same proxy instance. A comparison against a non-proxy is FALSE
            %   (mirrors Python's `isinstance(other, ElementProxy)` guard).
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__eq__
            if ~isa(a, "mat2doc.shared.ElementProxy") || ...
                    ~isa(b, "mat2doc.shared.ElementProxy")
                tf = false;             % not both proxies -> not equal
                return
            end
            % `element_ == element_`: XmlElement's SEALED handle-identity eq
            % (design.md section 2) -- Python `self._element is other._element`.
            tf = (a.element_ == b.element_);
        end

        function tf = ne(a, b)
            % NE True iff a and b do NOT wrap the same oxml element (shared.py 301-304).
            %
            %   The complement of eq; a comparison against a non-proxy is TRUE
            %   (mirrors Python's `__ne__` returning True for a non-proxy other).
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::ElementProxy.__ne__
            if ~isa(a, "mat2doc.shared.ElementProxy") || ...
                    ~isa(b, "mat2doc.shared.ElementProxy")
                tf = true;              % not both proxies -> unequal
                return
            end
            tf = (a.element_ ~= b.element_);
        end
    end
end
