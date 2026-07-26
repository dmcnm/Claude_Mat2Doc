classdef SettingsPart < mat2doc.opc.XmlPart
% SETTINGSPART Document-level settings part of a WML (.docx) package.
%
%   A pure XmlPart shell (P2-2): it PARSES on load and RE-SERIALIZES on save
%   through serialize_part_xml (byte-matched to lxml). blob/element/part inherit
%   from XmlPart unchanged; SettingsPart adds a custom constructor (stores the
%   settings element), `default()` (+ its template) and the `settings` proxy
%   accessor. No oxml element classes are needed -- parse_xml yields generic
%   elements until P5-1 registers CT_Settings.
%
%   FLIP (byte-neutral): PartFactory maps WML_SETTINGS -> SettingsPart at P2-2.
%   SettingsPart inherits XmlPart.blob and its own static `load` constructs a
%   SettingsPart, so the reloaded part's TYPE changes but the emitted bytes are
%   IDENTICAL to the base-XmlPart dispatch (M1 17/17 L1 unchanged).
%
%   CUSTOM CONSTRUCTOR (settings.py 22-26): unlike the plain XmlPart, docx
%   SettingsPart.__init__ ALSO stores `self._settings = element`. Ported here as
%   settings_ (the private cache the P5-1 `settings` proxy will read). It is dead
%   state until P5-1, but the constructor is faithful.
%
%   INHERITED-STATIC TRAP: own `load` (opc/part.py 229-232, cls=SettingsPart).
%
%   P2-2 SCOPE: `settings` (the Settings proxy accessor) is a P5-1 FEATURE STUB.
%   `default()` (+ template) is ported but NOT on the M1 open/save path
%   (default.docx ships settings.xml -> loads via `load`).
%
%   ARG ORDER (docx): SettingsPart(partname, content_type, element, package)
%   (settings.py 22-24).
%
%   Ported from python-docx v1.2.0: src/docx/parts/settings.py::SettingsPart
%   (__init__ 22-26, default 28-34 + _default_settings_xml 44-50 LIVE;
%   settings 36-42 -> P5-1 stub).

    properties (Access = private)
        settings_        % self._settings = element (settings.py 26); P5-1 cache
    end

    methods
        function obj = SettingsPart(partname, content_type, element, package)
            % __init__ (settings.py 22-26): forward to XmlPart, then store the
            %   settings element. Pass-through for the parser/factory (design.md
            %   CT_*/part constructor contract).
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
            obj.settings_ = element;   % Python: self._settings = element
        end

        function s = settings(obj) %#ok<MANU,STOUT>
            % SETTINGS STUB (settings.py 36-42, @property). Owner: P5-1.
            %   Faithful body: return Settings(self._settings). The Settings proxy
            %   lands at P5-1.
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.settings.Settings (owning WP: P5-1 settings tier) " + ...
                "required by mat2doc.parts.SettingsPart.settings");
        end
    end

    methods (Static)
        function obj = default(package)
            % DEFAULT (settings.py 28-34, @classmethod): a newly created settings
            %   part containing a default `w:settings` element tree. Python:
            %     partname = PackURI("/word/settings.xml")
            %     content_type = CT.WML_SETTINGS
            %     element = cast(CT_Settings, parse_xml(cls._default_settings_xml()))
            %     return cls(partname, content_type, element, package)
            partname = mat2doc.opc.PackURI("/word/settings.xml");
            content_type = mat2doc.opc.CONTENT_TYPE.WML_SETTINGS;
            element = mat2doc.oxml.parse_xml( ...
                mat2doc.parts.SettingsPart.default_settings_xml_());
            obj = mat2doc.parts.SettingsPart(partname, content_type, element, package);
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a SettingsPart (opc/part.py 229-232, cls=SettingsPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.SettingsPart(partname, content_type, element, package);
        end
    end

    methods (Static, Access = private)
        function xml_bytes = default_settings_xml_()
            % _default_settings_xml (settings.py 44-50, @classmethod): the bytes
            %   of the default settings part template, read BINARY from
            %   +mat2doc/templates/default-settings.xml (the Python
            %   os.path.split(__file__)[0] + ".." + templates analogue).
            partsdir = fileparts(mfilename("fullpath"));   % +mat2doc/+parts
            pkgdir = fileparts(partsdir);                  % +mat2doc
            path = fullfile(pkgdir, "templates", "default-settings.xml");
            fid = fopen(path, "rb");
            if fid == -1
                error("mat2doc:FileNotFoundError", ...
                    "default settings template not found: %s", path);
            end
            cleanup = onCleanup(@() fclose(fid)); %#ok<NASGU>
            xml_bytes = fread(fid, Inf, "*uint8")';
        end
    end
end
