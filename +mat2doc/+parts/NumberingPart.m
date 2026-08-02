classdef NumberingPart < mat2doc.opc.XmlPart
% NUMBERINGPART Proxy for the numbering.xml part (numbering definitions).
%
%   A pure XmlPart shell (P2-2): it PARSES on load and RE-SERIALIZES on save
%   through serialize_part_xml (byte-matched to lxml). blob/element/part inherit
%   from XmlPart unchanged. No oxml element classes are needed here -- parse_xml
%   yields generic elements until P8-1 registers CT_Numbering.
%
%   FLIP (byte-neutral): PartFactory maps WML_NUMBERING -> NumberingPart at P2-2.
%   NumberingPart inherits XmlPart.blob and its own static `load` constructs a
%   NumberingPart, so the reloaded part's TYPE changes but the emitted bytes are
%   IDENTICAL to the base-XmlPart dispatch (M1 17/17 L1 unchanged; default.docx
%   ships numbering.xml -> loads here).
%
%   INHERITED-STATIC TRAP: own `load` (opc/part.py 229-232, cls=NumberingPart)
%   so the flip actually lands on NumberingPart.
%
%   `new()` -- FAITHFUL NotImplementedError, NOT a port stub: python-docx v1.2.0
%   itself declares `def new(cls): raise NotImplementedError` (numbering.py
%   11-14). Reproduced verbatim as mat2doc:NotImplementedError (design.md section
%   2 exception mapping). It is unreached at P2-2 (DocumentPart.numbering_part,
%   the only caller of new(), is itself a P8-1 feature stub).
%
%   `numbering_definitions` -- UN-STUBBED at P8-1 (@lazyproperty): builds a
%   NumberingDefinitions_ over self._element (the parsed <w:numbering> ->
%   CT_Numbering). Cached via a logical flag (design.md @lazyproperty rule; NEVER
%   isempty as the sentinel). Never on the open/save path.
%
%   ARG ORDER (docx): NumberingPart(partname, content_type, element, package).
%
%   Ported from python-docx v1.2.0: src/docx/parts/numbering.py::NumberingPart
%   (new 11-14 faithful NotImplementedError; numbering_definitions 16-20 +
%   _NumberingDefinitions 23-32 LIVE at P8-1.)

    properties (Access = private)
        numbering_definitions_cache_                       % numbering_definitions lazyproperty cache
        numbering_definitions_computed_ (1,1) logical = false
    end

    methods
        function obj = NumberingPart(partname, content_type, element, package)
            % Pass-through to the XmlPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
        end

        function nd = numbering_definitions(obj)
            % NUMBERING_DEFINITIONS (numbering.py 16-20, @lazyproperty): the
            %   NumberingDefinitions_ instance containing the numbering definitions
            %   (<w:num> proxies) for this part. Python:
            %     return _NumberingDefinitions(self._element)
            %   Cached via a logical flag (design.md @lazyproperty rule; NEVER
            %   isempty as the sentinel).
            if ~obj.numbering_definitions_computed_
                obj.numbering_definitions_cache_ = ...
                    mat2doc.parts.NumberingDefinitions_(obj.element());
                obj.numbering_definitions_computed_ = true;
            end
            nd = obj.numbering_definitions_cache_;
        end
    end

    methods (Static)
        function obj = new() %#ok<STOUT>
            % NEW (numbering.py 11-14, @classmethod): FAITHFUL NotImplementedError
            %   -- python-docx v1.2.0 itself leaves this unimplemented (`raise
            %   NotImplementedError`). This is the upstream behavior, NOT a
            %   mat2doc:notYetPorted stub.
            error("mat2doc:NotImplementedError", "%s", ...
                "mat2doc.parts.NumberingPart.new mirrors python-docx v1.2.0, " + ...
                "which raises NotImplementedError (numbering.py 11-14)");
        end

        function obj = load(partname, content_type, blob, package)
            % LOAD OWN static override (inherited-static trap): parse the blob and
            %   construct a NumberingPart (opc/part.py 229-232, cls=NumberingPart).
            element = mat2doc.oxml.parse_xml(blob);
            obj = mat2doc.parts.NumberingPart(partname, content_type, element, package);
        end
    end
end
