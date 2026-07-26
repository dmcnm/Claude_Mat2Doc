classdef SerializedRelationships_ < handle
% SERIALIZEDRELATIONSHIPS_ Read-only sequence of SerializedRelationship_ objects.
%
%   Corresponds to the relationships item XML (.rels) passed to load_from_xml. The
%   Python original supports iteration via __iter__; MATLAB exposes to_array()
%   (design.md section 2 dunder mapping: `for x in srels` -> `for x =
%   srels.to_array()`). A handle class (the Python object is built by appending to
%   an internal list, and PackageReader holds references to it).
%
%   UNDERSCORE ROTATION (design.md section 2): `_SerializedRelationships` ->
%   SerializedRelationships_.
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgreader.py::_SerializedRelationships
%   (lines 230-254)

    properties (Access = private)
        srels_ = mat2doc.opc.SerializedRelationship_.empty(1, 0)  % 1xN handle array
    end

    methods
        function obj = SerializedRelationships_()
            % pkgreader.py 234-236: start with an empty list.
        end

        function arr = to_array(obj)
            % __iter__ (pkgreader.py 238-240): iterate self._srels. Returns the
            %   1xN SerializedRelationship_ array in document (append) order.
            arr = obj.srels_;
        end
    end

    methods (Static)
        function srels = load_from_xml(baseURI, rels_item_xml)
            % LOAD_FROM_XML Build a SerializedRelationships_ from .rels bytes
            %   (pkgreader.py 242-254). Returns an EMPTY collection when
            %   rels_item_xml is None ([], H3).
            %
            %   V-3: this loader lives HERE (pkgreader.py), not in rel.py -- it is
            %   the serialized-relationship unmarshaller that P1-5's brief
            %   mis-located. Faithful to v1.2.0.
            srels = mat2doc.opc.SerializedRelationships_();
            % Python: `if rels_item_xml is not None`. The phys reader returns []
            % (None) when the .rels item is absent; a present .rels blob is
            % non-empty uint8, so ~isempty is an exact None-test here (H3).
            if ~isempty(rels_item_xml)
                rels_elm = mat2doc.opc.oxml.parse_xml(rels_item_xml);
                rel_elms = rels_elm.Relationship_lst;
                for k = 1:numel(rel_elms)
                    srels.srels_(end + 1) = ...
                        mat2doc.opc.SerializedRelationship_(baseURI, rel_elms(k));
                end
            end
        end
    end
end
