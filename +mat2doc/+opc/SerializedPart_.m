classdef SerializedPart_ < handle
% SERIALIZEDPART_ Value object for an OPC package part (serialized form).
%
%   Provides read-only access to the partname, content type, referring
%   relationship type, blob, and serialized relationships for a part read from a
%   physical package. "Serialized" means targets are referred to by partname, not
%   by an in-memory Part object. A handle class because the Python original is a
%   plain object with identity (the PackageReader collects a tuple of them and
%   iterates them by reference).
%
%   UNDERSCORE ROTATION (design.md section 2): `_SerializedPart` -> SerializedPart_.
%
%   Ported from python-docx v1.2.0: src/docx/opc/pkgreader.py::_SerializedPart
%   (lines 130-164)

    properties (Access = private)
        partname_       % PackURI
        content_type_   % string
        reltype_        % string (referring relationship type)
        blob_           % uint8 part bytes
        srels_          % SerializedRelationships_
    end

    properties (Dependent, SetAccess = private)
        partname
        content_type
        blob
        reltype
        srels
    end

    methods
        function obj = SerializedPart_(partname, content_type, reltype, blob, srels)
            % pkgreader.py 137-143: store the five values verbatim.
            obj.partname_ = partname;
            obj.content_type_ = content_type;
            obj.reltype_ = reltype;
            obj.blob_ = blob;
            obj.srels_ = srels;
        end

        function value = get.partname(obj)
            % @property partname (pkgreader.py 145-147).
            value = obj.partname_;
        end

        function value = get.content_type(obj)
            % @property content_type (pkgreader.py 149-151).
            value = obj.content_type_;
        end

        function value = get.blob(obj)
            % @property blob (pkgreader.py 153-155).
            value = obj.blob_;
        end

        function value = get.reltype(obj)
            % @property reltype (pkgreader.py 157-160): referring relationship type.
            value = obj.reltype_;
        end

        function value = get.srels(obj)
            % @property srels (pkgreader.py 162-164).
            value = obj.srels_;
        end
    end
end
