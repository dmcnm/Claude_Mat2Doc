classdef BabelFish
% BABELFISH Translate style names between UI form and internal/styles.xml form.
%
%   Translates special-case style names from UI name (e.g. "Heading 1") to
%   internal/styles.xml name (e.g. "heading 1") and back. Names not in the fixed
%   alias table pass through unchanged (both directions).
%
%   H15 (case sensitivity): the alias lookups are EXACT, case-sensitive string
%   matches (Python dict.get keyed on the exact string), so "Heading 1" and
%   "heading 1" are distinct keys and only the listed pairs are translated. No
%   lower()/upper() is applied anywhere -- an unknown name is returned verbatim.
%
%   STORAGE (H11): the 12 alias pairs (styles/__init__.py 12-25, VERBATIM) are
%   held as a Constant Nx2 string array ALIASES = [ui, internal]. Python builds
%   two dicts from the same tuple: internal_style_names = dict(style_aliases)
%   maps UI->internal (used by ui2internal); ui_style_names = {internal: ui}
%   maps internal->UI (used by internal2ui). Both are reproduced here by scanning
%   column 1 (ui2internal) or column 2 (internal2ui) of ALIASES. The pairs are a
%   bijection, so no key-collision ambiguity arises.
%
%   NOTE (module-vs-class flattening, FLAG-3): BabelFish lives in Python's
%   `docx.styles` package `__init__`; StyleFactory/BaseStyle/... live in
%   `docx.styles.style` and Styles in `docx.styles.styles`. The Mat2Doc port
%   FLATTENS all of these into the single `+mat2doc\+styles` package (no
%   module-mirroring sub-packages), so this class is `mat2doc.styles.BabelFish`.
%
%   Example:
%       mat2doc.styles.BabelFish.ui2internal("Heading 1")   % "heading 1"
%       mat2doc.styles.BabelFish.internal2ui("heading 1")   % "Heading 1"
%       mat2doc.styles.BabelFish.ui2internal("Normal")      % "Normal" (passthrough)
%
%   Ported from python-docx v1.2.0: src/docx/styles/__init__.py::BabelFish

    properties (Constant, Hidden)
        % style_aliases (styles/__init__.py 12-25), VERBATIM as [UI, internal] pairs.
        ALIASES = [ ...
            "Caption",   "caption"; ...
            "Footer",    "footer"; ...
            "Header",    "header"; ...
            "Heading 1", "heading 1"; ...
            "Heading 2", "heading 2"; ...
            "Heading 3", "heading 3"; ...
            "Heading 4", "heading 4"; ...
            "Heading 5", "heading 5"; ...
            "Heading 6", "heading 6"; ...
            "Heading 7", "heading 7"; ...
            "Heading 8", "heading 8"; ...
            "Heading 9", "heading 9" ]
    end

    methods (Static)
        function internal_style_name = ui2internal(ui_style_name)
            % UI2INTERNAL Internal style name for `ui_style_name`, else passthrough.
            %   Python (styles/__init__.py 30-34):
            %     return cls.internal_style_names.get(ui_style_name, ui_style_name)
            %   internal_style_names maps UI->internal. On a miss, the arg is
            %   returned unchanged.
            %
            %   Ported from python-docx v1.2.0: styles/__init__.py::BabelFish.ui2internal
            aliases = mat2doc.styles.BabelFish.ALIASES;
            idx = find(aliases(:, 1) == ui_style_name, 1);   % H15: exact, case-sensitive
            if isempty(idx)   % Python dict.get miss -> default (the arg itself)
                internal_style_name = ui_style_name;
                return
            end
            internal_style_name = aliases(idx, 2);
        end

        function ui_style_name = internal2ui(internal_style_name)
            % INTERNAL2UI UI style name for `internal_style_name`, else passthrough.
            %   Python (styles/__init__.py 36-40):
            %     return cls.ui_style_names.get(internal_style_name, internal_style_name)
            %   ui_style_names maps internal->UI. On a miss, the arg is returned
            %   unchanged.
            %
            %   Ported from python-docx v1.2.0: styles/__init__.py::BabelFish.internal2ui
            aliases = mat2doc.styles.BabelFish.ALIASES;
            idx = find(aliases(:, 2) == internal_style_name, 1);   % H15: exact, case-sensitive
            if isempty(idx)   % Python dict.get miss -> default (the arg itself)
                ui_style_name = internal_style_name;
                return
            end
            ui_style_name = aliases(idx, 1);
        end
    end
end
