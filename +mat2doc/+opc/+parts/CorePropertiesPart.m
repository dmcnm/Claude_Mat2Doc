classdef CorePropertiesPart < mat2doc.opc.XmlPart
% COREPROPERTIESPART Corresponds to the part named /docProps/core.xml.
%
%   The "core" is short for "Dublin Core" and contains document metadata
%   relatively common across documents of all types, not just DOCX. Subclasses
%   XmlPart -- so it PARSES on load and RE-SERIALIZES on save through
%   serialize_part_xml (byte-matched to lxml) -- and wraps a
%   mat2doc.oxml.coreprops.CT_CoreProperties element, exposed through a
%   CoreProperties proxy (core_properties).
%
%   BYTE-IDENTITY (M1 round-trip): CorePropertiesPart inherits XmlPart.blob
%   unchanged (parse + serialize_part_xml) and does not mutate the parsed tree
%   on a plain open+save, so registering it for OPC_CORE_PROPERTIES would produce
%   bytes IDENTICAL to the base-XmlPart dispatch. The docProps/core.xml of the
%   bundled default.docx already carries its <cp:revision>/<dcterms:*> values and
%   the xsi decl, so open+save NEVER re-stamps it.
%
%   M1 UNREACHABLE (both the default() path AND this subclass): at M1 the
%   PartFactory maps OPC_CORE_PROPERTIES to the BASE mat2doc.opc.XmlPart (the
%   XmlPart-vs-Part split; PartFactory.part_type_for_ row comment "P2:
%   CorePropertiesPart"), so core.xml loads as a base XmlPart and is
%   re-serialized verbatim -- this subclass is NOT instantiated at M1. It is
%   ported now (registry: CT_CoreProperties for the cp:coreProperties ROOT is
%   wired; the PART wiring is a P2 refinement). `default()` is additionally dead
%   for any round-trip of a package that HAS a core.xml (it only builds one from
%   scratch for a core-less package).
%
%   D-COREPROPS-TIME (adopt-only; deviation ledger, SIGNED-PROVISIONAL):
%   `default()` stamps the current wall-clock time into `modified` (a
%   datetime.now(UTC) analogue). A wall-clock deviation of the same family as
%   D-zip-time. DEAD for the byte-identity round-trip (default.docx HAS core.xml,
%   so `default()` is never called on it) and LIVE only when creating
%   core-properties on a core-less package or when a caller sets `modified`. See
%   validation\summary\proofs\D-coreprops-time_explained.md.
%
%   The read/write path delegates entirely to the wrapped CT_CoreProperties, so
%   its DEVIATIONS apply transitively: `revision` reads through parse_int_ and
%   `created`/`modified`/`last_printed` through the W3CDTF grammar (both D-002,
%   ASCII-only), and the created/modified setters carry the xsi nsdecl hoist
%   (D-serializer-nsdecl). See the CT_CoreProperties header + deviation_ledger.md.
%
%   MATLAB does NOT inherit constructors or dispatch inherited static methods to
%   the subclass, so this class declares its own pass-through constructor and its
%   own `load` (the PartFactory entry point) constructing a CorePropertiesPart --
%   the faithful realization of Python's inherited-but-cls-bound XmlPart.load
%   (opc/part.py 229-232, where cls is CorePropertiesPart).
%
%   ARG ORDER (docx): CorePropertiesPart(partname, content_type, element,
%   package) -- element third, package last (opc/part.py XmlPart.__init__ order,
%   opc/parts/coreprops.py:47).
%
%   Example:
%       tpl = fullfile(fileparts(fileparts(which( ...
%           "mat2doc.opc.OpcPackage"))), "templates", "default.docx");
%       % (at P2, once OpcPackage.core_properties + PartFactory wiring land)
%
%   Ported from python-docx v1.2.0: src/docx/opc/parts/coreprops.py::CorePropertiesPart

    properties (Dependent)
        core_properties   % CoreProperties proxy over this part's element (parts/coreprops.py 37-41)
    end

    methods
        function obj = CorePropertiesPart(partname, content_type, element, package)
            % Pass-through to the XmlPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation. ARG
            %   ORDER element-before-package matches docx XmlPart.__init__.
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
        end

        function cp = get.core_properties(obj)
            % CORE_PROPERTIES A CoreProperties proxy providing read/write access to
            %   the core properties in this part (parts/coreprops.py 37-41):
            %   CoreProperties(self.element).
            cp = mat2doc.opc.CoreProperties(obj.element());
        end
    end

    methods (Static)
        function core_properties_part = default(package)
            % DEFAULT Return a default new CorePropertiesPart (parts/coreprops.py 25-35).
            %   A base for adding core-properties to a package that has none.
            %
            %   D-COREPROPS-TIME: the `modified` stamp below reads the wall clock
            %   (datetime.now(UTC)). SIGNED-PROVISIONAL deviation; live only on
            %   this create-a-default path (see class header + ledger).
            core_properties_part = mat2doc.opc.parts.CorePropertiesPart.new_(package);
            core_properties = core_properties_part.core_properties;
            core_properties.title = "Word Document";
            core_properties.last_modified_by = "python-docx";
            core_properties.revision = 1;
            % dt.datetime.now(dt.timezone.utc): UTC wall clock. Cleared to a naive
            % datetime (same UTC fields) and floored to the whole second so the
            % strftime("%Y-%m-%dT%H:%M:%SZ") stamp truncates identically to Python.
            now_utc = datetime("now", "TimeZone", "UTC");
            now_utc.TimeZone = "";
            core_properties.modified = dateshift(now_utc, "start", "second");
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD PartFactory entry point -- the faithful MATLAB realization of
            %   the inherited XmlPart.load with cls=CorePropertiesPart (opc/part.py
            %   229-232): parse the blob into a CT_CoreProperties tree and
            %   construct a CorePropertiesPart. ARG ORDER blob-before-package
            %   matches PartFactory. (Not wired into PartFactory until P2.)
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.opc.parts.CorePropertiesPart( ...
                partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function obj = new_(package)
            % NEW_ Return a new empty CorePropertiesPart (parts/coreprops.py 43-48).
            %   ARG ORDER: element (CT_CoreProperties.new()) before package.
            obj = mat2doc.opc.parts.CorePropertiesPart( ...
                mat2doc.opc.PackURI("/docProps/core.xml"), ...
                mat2doc.opc.CONTENT_TYPE.OPC_CORE_PROPERTIES, ...
                mat2doc.oxml.coreprops.CT_CoreProperties.new(), ...
                package);
        end
    end
end
