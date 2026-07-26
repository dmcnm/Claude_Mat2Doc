classdef PackURI
% PACKURI Proxy for a pack URI (partname). Behaves as a string otherwise.
%
%   uri = MAT2DOC.OPC.PACKURI(pack_uri_str) where pack_uri_str MUST begin with
%   "/", e.g. "/word/document.xml". Provides the baseURI / ext / filename /
%   idx / membername / rels_uri utility properties and the from_rel_ref /
%   relative_ref helpers.
%
%   Mapping note (H2): the Python class subclasses `str`. MATLAB classdef cannot
%   subclass string, so PackURI is a VALUE class wrapping the URI text in a
%   private property, with explicit string()/char() conversions and eq/ne so
%   partnames compare and index like the Python str. Any Python call site that
%   uses a PackURI directly AS a str (dict key, formatting, path join) maps to
%   string(obj) here.
%
%   Immutability: Python str is immutable, so this is a VALUE class (copies are
%   indistinguishable), matching the value-object semantics.
%
%   REGEX note (H12): docx's `_filename_re` is "([a-zA-Z]+)([1-9][0-9]*)?"
%   (packuri.py:18) -- the trailing-int group is [1-9][0-9]* (NO leading zero),
%   which DIFFERS from python-pptx's "([0-9][0-9]*)?" (Mat2Ppt used [0-9]). Ported
%   the docx form exactly: a filename whose numeric suffix begins with '0' (e.g.
%   "slide01") yields idx = [] (None), not 1.
%
%   Properties (read-only, mirror Python @property):
%       baseURI    - directory portion, e.g. "/word" for "/word/document.xml";
%                    "/" for the package pseudo-partname "/"
%       ext        - extension without the leading period, e.g. "xml" ("" none)
%       filename   - final component, e.g. "document.xml" ("" for "/")
%       idx        - int "tuple" partname index (e.g. 21) or [] (None) singleton
%       membername - the URI without the leading slash (the Zip membername);
%                    "" for "/"
%       rels_uri   - PackURI of the .rels part for this partname
%
%   Example:
%       uri = mat2doc.opc.PackURI("/word/slides/slide1.xml");
%       disp(uri.baseURI)              % "/word/slides"
%       disp(uri.membername)           % "word/slides/slide1.xml"  (Zip key)
%       disp(uri.idx)                  % 1  (the "tuple" partname index)
%       disp(string(uri.rels_uri))     % "/word/slides/_rels/slide1.xml.rels"
%
%   Ported from python-docx v1.2.0: src/docx/opc/packuri.py::PackURI

    properties (Access = private)
        uri_ (1,1) string = "/"   % the pack-URI text (the str value in Python)
    end

    properties (Dependent, SetAccess = private)
        baseURI
        ext
        filename
        idx
        membername
        rels_uri
    end

    methods
        function obj = PackURI(pack_uri_str)
            % Ported from packuri.py::PackURI.__new__ (lines 20-24).
            arguments
                pack_uri_str (1,1) string
            end
            c = char(pack_uri_str);
            % Python: if pack_uri_str[0] != "/": raise ValueError. On an empty
            % string, Python str[0] raises IndexError; PackURI is never
            % constructed empty, but guard to avoid an out-of-range char index.
            if isempty(c) || c(1) ~= '/'
                error("mat2doc:ValueError", ...
                    "PackURI must begin with slash, got '%s'", pack_uri_str);
            end
            obj.uri_ = pack_uri_str;
        end

        function value = get.baseURI(obj)
            % packuri.py baseURI (lines 33-40): posixpath.split(self)[0].
            [head, ~] = mat2doc.opc.PackURI.ppSplit_(obj.uri_);
            value = head;
        end

        function value = get.ext(obj)
            % packuri.py ext (lines 42-50): splitext[1] w/o leading period.
            [~, raw_ext] = mat2doc.opc.PackURI.ppSplitext_(obj.uri_);
            if startsWith(raw_ext, ".")
                value = extractAfter(raw_ext, 1);
            else
                value = raw_ext;
            end
        end

        function value = get.filename(obj)
            % packuri.py filename (lines 52-59): posixpath.split(self)[1].
            [~, tail] = mat2doc.opc.PackURI.ppSplit_(obj.uri_);
            value = tail;
        end

        function value = get.idx(obj)
            % packuri.py idx (lines 61-75): int "tuple" index or None.
            fn = obj.filename;
            if fn == ""
                value = [];   % Python: if not filename: return None (H3)
                return
            end
            [name_part, ~] = mat2doc.opc.PackURI.ppSplitext_(fn);
            % Python: self._filename_re = re.compile("([a-zA-Z]+)([1-9][0-9]*)?")
            % re.match anchors at the start (^); the trailing group is optional
            % and its first digit is 1-9 (no leading zero) -- see REGEX note.
            tok = regexp(name_part, '^([a-zA-Z]+)([1-9][0-9]*)?', 'tokens', 'once');
            if isempty(tok)
                value = [];   % Python: match is None -> return None
                return
            end
            if numel(tok) >= 2 && tok{2} ~= ""
                value = str2double(tok{2});   % Python: if match.group(2): int(...)
            else
                value = [];   % Python: return None
            end
        end

        function value = get.membername(obj)
            % packuri.py membername (lines 77-84): self[1:] (leading slash off).
            c = char(obj.uri_);
            value = string(c(2:end));   % "" for "/"
        end

        function value = get.rels_uri(obj)
            % packuri.py rels_uri (lines 96-105): join(baseURI, "_rels",
            %   "<filename>.rels").
            rels_filename = obj.filename + ".rels";
            rels_uri_str = mat2doc.opc.PackURI.ppJoin_( ...
                obj.baseURI, "_rels", rels_filename);
            value = mat2doc.opc.PackURI(rels_uri_str);
        end

        function ref = relative_ref(obj, baseURI)
            % RELATIVE_REF Relative reference to this item from baseURI
            %   (packuri.py 86-94). baseURI == "/" -> self[1:] (posixpath 2.6
            %   relpath-from-root workaround); else posixpath.relpath.
            arguments
                obj (1,1) mat2doc.opc.PackURI
                baseURI (1,1) string
            end
            if baseURI == "/"
                c = char(obj.uri_);
                ref = string(c(2:end));
            else
                ref = mat2doc.opc.PackURI.ppRelpath_(obj.uri_, baseURI);
            end
        end

        function s = string(obj)
            % STRING The pack-URI text (Python: the str value itself).
            s = obj.uri_;
        end

        function c = char(obj)
            c = char(obj.uri_);
        end

        function tf = eq(a, b)
            % EQ String-identity of the pack URIs (Python str __eq__). Accepts a
            %   PackURI or a string/char on either side, so partnames compare and
            %   serve as lookup keys like the Python str.
            tf = mat2doc.opc.PackURI.asText_(a) == mat2doc.opc.PackURI.asText_(b);
        end

        function tf = ne(a, b)
            tf = ~eq(a, b);
        end
    end

    methods (Static)
        function uri = from_rel_ref(baseURI, relative_ref)
            % FROM_REL_REF Absolute pack URI from translating relative_ref onto
            %   baseURI (packuri.py 26-31): abspath(join(baseURI, relative_ref)).
            arguments
                baseURI (1,1) string
                relative_ref (1,1) string
            end
            joined_uri = mat2doc.opc.PackURI.ppJoin_(baseURI, relative_ref);
            abs_uri = mat2doc.opc.PackURI.ppAbspath_(joined_uri);
            uri = mat2doc.opc.PackURI(abs_uri);
        end
    end

    % ------------------------------------------------------------------
    % posixpath helpers (pure-string ports of the posixpath functions the
    % Python PackURI uses -- split / splitext / join / abspath / relpath).
    % Pack URIs are always POSIX '/'-separated and (for the used surface)
    % absolute, so these mirror posixpath exactly on the inputs docx passes.
    % ------------------------------------------------------------------
    methods (Static, Access = private)
        function t = asText_(x)
            if isa(x, "mat2doc.opc.PackURI")
                t = x.uri_;
            else
                t = string(x);
            end
        end

        function [head, tail] = ppSplit_(p)
            % posixpath.split: (head, tail), tail the last component.
            c = char(p);
            i = find(c == '/', 1, 'last');
            if isempty(i)
                head = "";
                tail = string(c);
                return
            end
            headc = c(1:i);       % includes the trailing slash
            tail = string(c(i + 1:end));
            % Python: if head and head != '/'*len(head): head = head.rstrip('/')
            if ~isempty(headc) && ~all(headc == '/')
                headc = regexprep(headc, '/+$', '');
            end
            head = string(headc);
        end

        function [root, ext] = ppSplitext_(p)
            % posixpath.splitext (genericpath._splitext with sep='/', extsep='.').
            c = char(p);
            sep_index = find(c == '/', 1, 'last');
            if isempty(sep_index); sep_index = 0; end
            dot_index = find(c == '.', 1, 'last');
            if isempty(dot_index); dot_index = 0; end
            if dot_index > sep_index
                % Skip a run of leading dots in the filename component.
                filename_index = sep_index + 1;
                found = false;
                while filename_index < dot_index
                    if c(filename_index) ~= '.'
                        found = true;
                        break
                    end
                    filename_index = filename_index + 1;
                end
                if found
                    root = string(c(1:dot_index - 1));
                    ext = string(c(dot_index:end));
                    return
                end
            end
            root = string(c);
            ext = "";
        end

        function out = ppJoin_(a, varargin)
            % posixpath.join for the docx-used cases.
            path = char(a);
            for k = 1:numel(varargin)
                b = char(varargin{k});
                if ~isempty(b) && b(1) == '/'
                    path = b;
                elseif isempty(path) || path(end) == '/'
                    path = [path, b]; %#ok<AGROW>
                else
                    path = [path, '/', b]; %#ok<AGROW>
                end
            end
            out = string(path);
        end

        function out = ppAbspath_(p)
            % posixpath.abspath on an already-absolute pack URI = normpath.
            %   (docx only calls from_rel_ref with an absolute baseURI, so the
            %   joined path is always absolute -- cwd is never consulted.)
            out = mat2doc.opc.PackURI.ppNormpath_(p);
        end

        function out = ppNormpath_(p)
            % posixpath.normpath: collapse '', '.', '..' and duplicate slashes.
            c = char(p);
            if isempty(c)
                out = ".";
                return
            end
            initial_slash = c(1) == '/';
            comps = split(string(c), "/");
            newc = strings(1, 0);
            for k = 1:numel(comps)
                comp = comps(k);
                if comp == "" || comp == "."
                    continue
                end
                if comp ~= ".." || (~initial_slash && isempty(newc)) || ...
                        (~isempty(newc) && newc(end) == "..")
                    newc(end + 1) = comp; %#ok<AGROW>
                elseif ~isempty(newc)
                    newc(end) = [];
                end
            end
            path = strjoin(newc, "/");
            if initial_slash
                path = "/" + path;
            end
            if path == ""
                out = ".";
            else
                out = path;
            end
        end

        function out = ppRelpath_(path, start)
            % posixpath.relpath(path, start) for absolute POSIX inputs.
            start_list = mat2doc.opc.PackURI.nonEmptyParts_( ...
                mat2doc.opc.PackURI.ppAbspath_(start));
            path_list = mat2doc.opc.PackURI.nonEmptyParts_( ...
                mat2doc.opc.PackURI.ppAbspath_(path));
            % i = len(commonprefix([start_list, path_list])) -- element-wise.
            i = 0;
            n = min(numel(start_list), numel(path_list));
            while i < n && start_list(i + 1) == path_list(i + 1)
                i = i + 1;
            end
            rel = [repmat("..", 1, numel(start_list) - i), path_list(i + 1:end)];
            if isempty(rel)
                out = ".";   % posixpath curdir
            else
                out = strjoin(rel, "/");
            end
        end

        function parts = nonEmptyParts_(p)
            comps = split(string(p), "/");
            parts = comps(comps ~= "")';
            if isempty(parts)
                parts = strings(1, 0);
            end
        end
    end
end
