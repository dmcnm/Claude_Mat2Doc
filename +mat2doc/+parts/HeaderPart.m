classdef HeaderPart < mat2doc.parts.StoryPart
% HEADERPART Definition of a section header (the word/headerN.xml part).
%
%   A StoryPart subclass (parts/hdrftr.py:36 `class HeaderPart(StoryPart)`): it
%   PARSES on load and RE-SERIALIZES on save through serialize_part_xml
%   (byte-matched to lxml), inheriting blob/element/paragraphs/add_paragraph from
%   StoryPart<XmlPart unchanged. HeaderPart adds only the `new` factory (+ its
%   template reader). This is the FIRST runtime-created part in Mat2Doc (headers
%   are separate parts with their own rels), landed at P5-3b.
%
%   FLIP (byte-neutral): PartFactory maps WML_HEADER -> HeaderPart at P5-3b.
%   Because HeaderPart inherits XmlPart.blob (parse + serialize_part_xml) and its
%   own static `load` constructs a HeaderPart, a reloaded header part's TYPE
%   changes but the emitted bytes are IDENTICAL to the previous base-XmlPart
%   dispatch. No M1 fixture loads a header part (default.docx has none), so the
%   flip is inert on the M1 sweep.
%
%   INHERITED-STATIC TRAP (DocumentPart.m:260 / StylesPart.m:65 precedent):
%   MATLAB does not dispatch an inherited static method to the calling subclass,
%   so HeaderPart declares its OWN `load` -- the faithful realization of Python's
%   inherited XmlPart.load with cls=HeaderPart (opc/part.py 229-232). Without it
%   the PartFactory flip would silently build a base XmlPart and the newly-created
%   header would round-trip through the wrong class.
%
%   ARG ORDER (docx): HeaderPart(partname, content_type, element, package).
%
%   Example:
%       hp = mat2doc.parts.HeaderPart.new(pkg);   % word/header1.xml (default template)
%       hp.paragraphs;                            % the header's paragraphs
%
%   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::HeaderPart
%   (new 39-45 + _default_header_xml 47-53).

    methods
        function obj = HeaderPart(partname, content_type, element, package)
            % Pass-through to the StoryPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.parts.StoryPart(partname, content_type, element, package);
        end
    end

    methods (Static)
        function obj = new(package)
            % NEW (hdrftr.py 39-45, @classmethod): a newly created header part.
            %   Python:
            %     partname = package.next_partname("/word/header%d.xml")
            %     content_type = CT.WML_HEADER
            %     element = parse_xml(cls._default_header_xml())
            %     return cls(partname, content_type, element, package)
            %   next_partname is LIVE (P1-6b); the "%d" number is DATA, not an
            %   index (H1-safe). parse_xml applies remove_blank_text to the parsed
            %   template, exactly as lxml does.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::HeaderPart.new
            partname = package.next_partname("/word/header%d.xml");
            content_type = mat2doc.opc.CONTENT_TYPE.WML_HEADER;
            element = mat2doc.oxml.parse_xml( ...
                mat2doc.parts.HeaderPart.default_header_xml_());
            obj = mat2doc.parts.HeaderPart(partname, content_type, element, package);
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a HeaderPart (opc/part.py 229-232, cls=HeaderPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.HeaderPart(partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function xml_bytes = default_header_xml_()
            % _default_header_xml (hdrftr.py 47-53, @classmethod): the bytes of the
            %   default header part template. Python reads
            %   os.path.join(os.path.split(__file__)[0], "..", "templates",
            %   "default-header.xml") in "rb" mode. MATLAB analogue: from this class
            %   file's directory (+mat2doc/+parts), go up one (..) to +mat2doc, then
            %   templates/default-header.xml. Read as BINARY (uint8) so the exact
            %   template bytes (incl. CRLF) reach parse_xml unchanged (the
            %   StylesPart._default_styles_xml precedent).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::HeaderPart._default_header_xml
            partsdir = fileparts(mfilename("fullpath"));   % +mat2doc/+parts
            pkgdir = fileparts(partsdir);                  % +mat2doc
            path = fullfile(pkgdir, "templates", "default-header.xml");
            fid = fopen(path, "rb");
            if fid == -1
                error("mat2doc:FileNotFoundError", ...
                    "default header template not found: %s", path);
            end
            cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
            xml_bytes = fread(fid, Inf, "*uint8")';
        end
    end
end
