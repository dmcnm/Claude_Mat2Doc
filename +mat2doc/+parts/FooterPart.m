classdef FooterPart < mat2doc.parts.StoryPart
% FOOTERPART Definition of a section footer (the word/footerN.xml part).
%
%   A StoryPart subclass (parts/hdrftr.py:16 `class FooterPart(StoryPart)`): it
%   PARSES on load and RE-SERIALIZES on save through serialize_part_xml
%   (byte-matched to lxml), inheriting blob/element/paragraphs/add_paragraph from
%   StoryPart<XmlPart unchanged. FooterPart adds only the `new` factory (+ its
%   template reader). A runtime-created separate part with its own rels, landed at
%   P5-3b.
%
%   FLIP (byte-neutral): PartFactory maps WML_FOOTER -> FooterPart at P5-3b.
%   Because FooterPart inherits XmlPart.blob (parse + serialize_part_xml) and its
%   own static `load` constructs a FooterPart, a reloaded footer part's TYPE
%   changes but the emitted bytes are IDENTICAL to the previous base-XmlPart
%   dispatch. No M1 fixture loads a footer part (default.docx has none), so the
%   flip is inert on the M1 sweep.
%
%   INHERITED-STATIC TRAP (DocumentPart.m:260 / StylesPart.m:65 precedent):
%   MATLAB does not dispatch an inherited static method to the calling subclass,
%   so FooterPart declares its OWN `load` -- the faithful realization of Python's
%   inherited XmlPart.load with cls=FooterPart (opc/part.py 229-232). Without it
%   the PartFactory flip would silently build a base XmlPart.
%
%   ARG ORDER (docx): FooterPart(partname, content_type, element, package).
%
%   Example:
%       fp = mat2doc.parts.FooterPart.new(pkg);   % word/footer1.xml (default template)
%       fp.paragraphs;                            % the footer's paragraphs
%
%   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::FooterPart
%   (new 19-25 + _default_footer_xml 27-33).

    methods
        function obj = FooterPart(partname, content_type, element, package)
            % Pass-through to the StoryPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.parts.StoryPart(partname, content_type, element, package);
        end
    end

    methods (Static)
        function obj = new(package)
            % NEW (hdrftr.py 19-25, @classmethod): a newly created footer part.
            %   Python:
            %     partname = package.next_partname("/word/footer%d.xml")
            %     content_type = CT.WML_FOOTER
            %     element = parse_xml(cls._default_footer_xml())
            %     return cls(partname, content_type, element, package)
            %   next_partname is LIVE (P1-6b); the "%d" number is DATA, not an
            %   index (H1-safe). parse_xml applies remove_blank_text to the parsed
            %   template, exactly as lxml does.
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::FooterPart.new
            partname = package.next_partname("/word/footer%d.xml");
            content_type = mat2doc.opc.CONTENT_TYPE.WML_FOOTER;
            element = mat2doc.oxml.parse_xml( ...
                mat2doc.parts.FooterPart.default_footer_xml_());
            obj = mat2doc.parts.FooterPart(partname, content_type, element, package);
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a FooterPart (opc/part.py 229-232, cls=FooterPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.FooterPart(partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function xml_bytes = default_footer_xml_()
            % _default_footer_xml (hdrftr.py 27-33, @classmethod): the bytes of the
            %   default footer part template. Python reads
            %   os.path.join(os.path.split(__file__)[0], "..", "templates",
            %   "default-footer.xml") in "rb" mode. MATLAB analogue: from this class
            %   file's directory (+mat2doc/+parts), go up one (..) to +mat2doc, then
            %   templates/default-footer.xml. Read as BINARY (uint8) so the exact
            %   template bytes (incl. CRLF) reach parse_xml unchanged (the
            %   StylesPart._default_styles_xml precedent).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/hdrftr.py::FooterPart._default_footer_xml
            partsdir = fileparts(mfilename("fullpath"));   % +mat2doc/+parts
            pkgdir = fileparts(partsdir);                  % +mat2doc
            path = fullfile(pkgdir, "templates", "default-footer.xml");
            fid = fopen(path, "rb");
            if fid == -1
                error("mat2doc:FileNotFoundError", ...
                    "default footer template not found: %s", path);
            end
            cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
            xml_bytes = fread(fid, Inf, "*uint8")';
        end
    end
end
