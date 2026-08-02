classdef CommentsPart < mat2doc.parts.StoryPart
% COMMENTSPART Container part for comments added to the document (word/comments.xml).
%
%   A StoryPart subclass (parts/comments.py:23 `class CommentsPart(StoryPart)`):
%   it PARSES on load and RE-SERIALIZES on save through serialize_part_xml
%   (byte-matched to lxml), inheriting blob/element/paragraphs/add_paragraph from
%   StoryPart<XmlPart. CommentsPart is the LAST new part type in the port. Unlike
%   the header/footer parts (numbered word/headerN.xml via next_partname),
%   comments live in the single fixed part word/comments.xml, created ON DEMAND:
%   default.docx ships NO comments part; DocumentPart._comments_part materializes
%   one via CommentsPart.default(package) on the first comment access (the
%   SettingsPart.default + header/footer separate-part precedents).
%
%   FLIP (byte-neutral): PartFactory maps WML_COMMENTS -> CommentsPart at P8-2.
%   CommentsPart inherits XmlPart.blob (parse + serialize_part_xml) and declares
%   its own static `load` constructing a CommentsPart, so a reloaded comments
%   part's TYPE changes but the emitted bytes are IDENTICAL to the prior
%   base-XmlPart dispatch. No M1 fixture loads a comments part (default.docx has
%   none), so the flip is inert on the M1 sweep -- M1 stays 17/17 L1.
%
%   CUSTOM CONSTRUCTOR (comments.py 26-30): unlike the plain StoryPart, docx
%   CommentsPart.__init__ ALSO stores `self._comments = element`. Ported here as
%   comments_ (the private cache the `comments` proxy accessor reads). Underscore
%   rotation (design.md section 2): _comments -> comments_.
%
%   INHERITED-STATIC TRAP (HeaderPart/SettingsPart precedent): MATLAB does not
%   dispatch an inherited static method to the calling subclass, so CommentsPart
%   declares its OWN `load` -- the faithful realization of Python's inherited
%   XmlPart.load with cls=CommentsPart (opc/part.py 229-232). Without it the
%   PartFactory flip would silently build a base XmlPart.
%
%   ARG ORDER (docx): CommentsPart(partname, content_type, element, package)
%   (comments.py 26-28) -- element third, package last.
%
%   Example:
%       cp = mat2doc.parts.CommentsPart.default(pkg);   % word/comments.xml (empty)
%       cs = cp.comments;                               % a Comments proxy
%       cs.add_comment("Hi", "Amy", "AJ");
%
%   Ported from python-docx v1.2.0: src/docx/parts/comments.py::CommentsPart
%   (__init__ 26-30, comments 32-35, default 37-43 + _default_comments_xml 45-51).

    properties (Access = private)
        comments_        % self._comments = element (comments.py 30): the CT_Comments
    end

    methods
        function obj = CommentsPart(partname, content_type, element, package)
            % __init__ (comments.py 26-30): forward to StoryPart, then store the
            %   comments element. Pass-through for the parser/factory (design.md
            %   CT_*/part constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.parts.StoryPart(partname, content_type, element, package);
            obj.comments_ = element;   % Python: self._comments = element
        end

        function c = comments(obj)
            % COMMENTS A Comments proxy for the `<w:comments>` root element of this
            %   part (comments.py 32-35, @property). Python: return
            %   Comments(self._comments, self). H5: minted fresh each access
            %   (python-docx does not cache the Comments object).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/comments.py::CommentsPart.comments
            c = mat2doc.comments.Comments(obj.comments_, obj);
        end
    end

    methods (Static)
        function obj = default(package)
            % DEFAULT (comments.py 37-43, @classmethod): a newly created comments
            %   part containing a default empty `<w:comments>` element. Python:
            %     partname = PackURI("/word/comments.xml")
            %     content_type = CT.WML_COMMENTS
            %     element = cast(CT_Comments, parse_xml(cls._default_comments_xml()))
            %     return cls(partname, content_type, element, package)
            %   parse_xml applies remove_blank_text to the parsed template (as lxml
            %   does). Registry -> CT_Comments for the <w:comments> root.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/comments.py::CommentsPart.default
            partname = mat2doc.opc.PackURI("/word/comments.xml");
            content_type = mat2doc.opc.CONTENT_TYPE.WML_COMMENTS;
            element = mat2doc.oxml.parse_xml( ...
                mat2doc.parts.CommentsPart.default_comments_xml_());
            obj = mat2doc.parts.CommentsPart(partname, content_type, element, package);
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a CommentsPart (opc/part.py 229-232, cls=CommentsPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.CommentsPart(partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function xml_bytes = default_comments_xml_()
            % _default_comments_xml (comments.py 45-51, @classmethod): the bytes of
            %   the default comments part template, read BINARY from
            %   +mat2doc/templates/default-comments.xml (the Python
            %   os.path.split(__file__)[0] + ".." + templates analogue). Read as
            %   uint8 so the exact template bytes (incl. CRLF) reach parse_xml
            %   unchanged (the StylesPart/HeaderPart _default_*_xml precedent).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/comments.py::CommentsPart._default_comments_xml
            partsdir = fileparts(mfilename("fullpath"));   % +mat2doc/+parts
            pkgdir = fileparts(partsdir);                  % +mat2doc
            path = fullfile(pkgdir, "templates", "default-comments.xml");
            fid = fopen(path, "rb");
            if fid == -1
                error("mat2doc:FileNotFoundError", ...
                    "default comments template not found: %s", path);
            end
            cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
            xml_bytes = fread(fid, Inf, "*uint8")';
        end
    end
end
