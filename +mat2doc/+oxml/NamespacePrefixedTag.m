classdef NamespacePrefixedTag
% NAMESPACEPREFIXEDTAG Value object that knows the semantics of an XML tag having a namespace prefix.
%
%   obj = MAT2DOC.OXML.NAMESPACEPREFIXEDTAG(nstag) where nstag is e.g. "w:p".
%
%   Mapping note (H2): the Python class subclasses `str`. MATLAB classdef
%   cannot subclass string, so this class wraps the tag text in a private
%   property and provides explicit string()/char() conversions. Any Python
%   call site that uses a NamespacePrefixedTag directly AS a str (formatting,
%   comparison, dict key) must call string(obj) explicitly when ported.
%
%   Immutability: Python str instances are immutable, so this is a MATLAB
%   VALUE class (not handle) -- copies are indistinguishable from the
%   original, matching Python value-object semantics.
%
%   Properties (read-only, mirror Python @property):
%       clark_name - "{uri}local" Clark notation (string)
%       local_part - local part of the tag, e.g. "foobar" for "f:foobar"
%       nsmap      - 1x1 struct with single field <prefix> -> URI
%       nspfx      - namespace prefix, e.g. "f" for "f:foobar"
%       nsuri      - namespace URI for the tag's prefix
%
%   Conversions (recover str-ness, H2):
%       string(obj) / char(obj) - the full prefixed-tag text (e.g. "w:p").
%       Use at any ported call site where Python uses the instance AS a str.
%
%   Inputs:  nstag - string scalar, a prefixed tag, e.g. "w:p"
%   Outputs: obj   - scalar NamespacePrefixedTag value object
%
%   Example:
%       nsptag = mat2doc.oxml.NamespacePrefixedTag("w:p");
%       nsptag.clark_name   % "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
%       nsptag.local_part   % "p"
%       string(nsptag)      % "w:p"
%       t2 = mat2doc.oxml.NamespacePrefixedTag.from_clark_name(nsptag.clark_name);
%       string(t2)          % "w:p"  (round-trip)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/ns.py::NamespacePrefixedTag

    properties (Access = private)
        nstag_ (1,1) string       % the full prefixed tag (the str value in Python)
        pfx_ (1,1) string         % Python self._pfx
        local_part_ (1,1) string  % Python self._local_part
        ns_uri_ (1,1) string      % Python self._ns_uri
    end

    properties (Dependent, SetAccess = private)
        clark_name
        local_part
        nsmap
        nspfx
        nsuri
    end

    methods
        function obj = NamespacePrefixedTag(nstag)
            % Ported from ns.py::NamespacePrefixedTag.__new__/__init__ (lines 32-37)
            arguments
                nstag (1,1) string
            end
            % Python: self._pfx, self._local_part = nstag.split(":")  (line 36)
            parts = split(nstag, ":");
            if numel(parts) < 2
                % CPython implicit ValueError on tuple unpack
                error("mat2doc:ValueError", ...
                    "not enough values to unpack (expected 2, got %d)", numel(parts));
            elseif numel(parts) > 2
                error("mat2doc:ValueError", "too many values to unpack (expected 2)");
            end
            obj.pfx_ = parts(1);
            obj.local_part_ = parts(2);
            % Python: self._ns_uri = nsmap[self._pfx]  (line 37)
            map = mat2doc.oxml.nsmap();
            if ~isfield(map, obj.pfx_)
                error("mat2doc:KeyError", "'%s'", obj.pfx_);
            end
            obj.ns_uri_ = map.(obj.pfx_);
            obj.nstag_ = nstag;
        end

        function value = get.clark_name(obj)
            % Ported from ns.py::NamespacePrefixedTag.clark_name (lines 39-41)
            value = "{" + obj.ns_uri_ + "}" + obj.local_part_;
        end

        function value = get.local_part(obj)
            % Ported from ns.py::NamespacePrefixedTag.local_part (lines 49-55)
            value = obj.local_part_;
        end

        function value = get.nsmap(obj)
            % Ported from ns.py::NamespacePrefixedTag.nsmap (lines 57-64)
            % Returns a 1x1 struct having a single field, mapping the namespace
            % prefix of this tag to its namespace name.
            value = struct();
            value.(obj.pfx_) = obj.ns_uri_;
        end

        function value = get.nspfx(obj)
            % Ported from ns.py::NamespacePrefixedTag.nspfx (lines 66-72)
            value = obj.pfx_;
        end

        function value = get.nsuri(obj)
            % Ported from ns.py::NamespacePrefixedTag.nsuri (lines 74-81)
            value = obj.ns_uri_;
        end

        function s = string(obj)
            % STRING The full prefixed-tag text (Python: the str value itself).
            s = obj.nstag_;
        end

        function c = char(obj)
            % CHAR The full prefixed-tag text as char (Python: the str value itself).
            c = char(obj.nstag_);
        end
    end

    methods (Static)
        function nsptag = from_clark_name(clark_name)
            % FROM_CLARK_NAME Construct from Clark notation, e.g. "{uri}local".
            %
            % Ported from ns.py::NamespacePrefixedTag.from_clark_name (lines 43-47)
            arguments
                clark_name (1,1) string
            end
            % Python: nsuri, local_name = clark_name[1:].split("}")  (line 45)
            % clark_name[1:] -> chars(2:end): Python slice tolerates short/empty
            % input (returns ""), extractAfter would error -- char slice matches. % IDX
            chars = char(clark_name);
            rest = string(chars(2:end));
            parts = split(rest, "}");
            if numel(parts) < 2
                error("mat2doc:ValueError", ...
                    "not enough values to unpack (expected 2, got %d)", numel(parts));
            elseif numel(parts) > 2
                error("mat2doc:ValueError", "too many values to unpack (expected 2)");
            end
            uri = parts(1);
            local_name = parts(2);
            % Python: nstag = "%s:%s" % (pfxmap[nsuri], local_name)  (line 46)
            map = mat2doc.oxml.pfxmap();
            row = find(map(:, 1) == uri, 1);
            if isempty(row)
                error("mat2doc:KeyError", "'%s'", uri);
            end
            nstag = map(row, 2) + ":" + local_name;
            nsptag = mat2doc.oxml.NamespacePrefixedTag(nstag);
        end
    end
end
