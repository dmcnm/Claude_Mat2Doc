classdef StylesPart < mat2doc.opc.XmlPart
% STYLESPART Proxy for the styles.xml part (style definitions for a document).
%
%   A pure XmlPart shell (P2-2): it PARSES on load and RE-SERIALIZES on save
%   through serialize_part_xml (byte-matched to lxml). blob/element/part inherit
%   from XmlPart unchanged; StylesPart adds only `default()` (+ its template) and
%   the `styles` proxy accessor. No oxml element classes are needed here --
%   parse_xml yields generic elements until P4-6 registers CT_Styles.
%
%   FLIP (byte-neutral): PartFactory maps WML_STYLES -> StylesPart at P2-2. Since
%   StylesPart inherits XmlPart.blob (parse + serialize_part_xml) and its own
%   static `load` constructs a StylesPart, the reloaded part's TYPE changes but
%   the emitted bytes are IDENTICAL to the previous base-XmlPart dispatch. The M1
%   17/17 L1 sweep is unchanged (styles.xml is byte-stable).
%
%   INHERITED-STATIC TRAP: MATLAB does not dispatch an inherited static method to
%   the calling subclass, so StylesPart declares its OWN `load` -- the faithful
%   realization of Python's inherited XmlPart.load with cls=StylesPart
%   (opc/part.py 229-232). Without it, PartFactory would build a base XmlPart and
%   the flip would be inert.
%
%   P2-2 SCOPE: `styles` (the Styles proxy accessor) is a P4-7 FEATURE STUB.
%   `default()` (+ its template) IS ported but is NOT on the M1 open/save path
%   (default.docx already ships styles.xml, so it loads via `load`; `default()`
%   builds a fresh part only for a styles-less package).
%
%   ARG ORDER (docx): StylesPart(partname, content_type, element, package).
%
%   Ported from python-docx v1.2.0: src/docx/parts/styles.py::StylesPart
%   (default 22-28 + _default_styles_xml 36-42 LIVE; styles 30-34 -> P4-7 stub).

    methods
        function obj = StylesPart(partname, content_type, element, package)
            % Pass-through to the XmlPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
        end

        function s = styles(obj)
            % STYLES The styles defined in this document (styles.py 30-34, @property).
            %   Python: return Styles(self.element). UN-STUBBED at P4-7a: returns a
            %   real mat2doc.styles.Styles proxy over this part's CT_Styles root.
            %   `element` is inherited from XmlPart (the parsed <w:styles>).
            %
            %   Ported from python-docx v1.2.0: src/docx/parts/styles.py::StylesPart.styles
            s = mat2doc.styles.Styles(obj.element());
        end
    end

    methods (Static)
        function obj = default(package)
            % DEFAULT (styles.py 22-28, @classmethod): a newly created styles part
            %   containing a default set of elements. Python:
            %     partname = PackURI("/word/styles.xml")
            %     content_type = CT.WML_STYLES
            %     element = parse_xml(cls._default_styles_xml())
            %     return cls(partname, content_type, element, package)
            partname = mat2doc.opc.PackURI("/word/styles.xml");
            content_type = mat2doc.opc.CONTENT_TYPE.WML_STYLES;
            element = mat2doc.oxml.parse_xml( ...
                mat2doc.parts.StylesPart.default_styles_xml_());
            obj = mat2doc.parts.StylesPart(partname, content_type, element, package);
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a StylesPart (opc/part.py 229-232, cls=StylesPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.StylesPart(partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function xml_bytes = default_styles_xml_()
            % _default_styles_xml (styles.py 36-42, @classmethod): the bytes of
            %   the default styles part template. Python reads
            %   os.path.join(os.path.split(__file__)[0], "..", "templates",
            %   "default-styles.xml") in "rb" mode. MATLAB analogue: from this
            %   class file's directory (+mat2doc/+parts), go up one (..) to
            %   +mat2doc, then templates/default-styles.xml. Read as BINARY (uint8)
            %   so the exact template bytes (incl. CRLF) round-trip.
            partsdir = fileparts(mfilename("fullpath"));   % +mat2doc/+parts
            pkgdir = fileparts(partsdir);                  % +mat2doc
            path = fullfile(pkgdir, "templates", "default-styles.xml");
            fid = fopen(path, "rb");
            if fid == -1
                error("mat2doc:FileNotFoundError", ...
                    "default styles template not found: %s", path);
            end
            cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
            xml_bytes = fread(fid, Inf, "*uint8")';
        end
    end
end
