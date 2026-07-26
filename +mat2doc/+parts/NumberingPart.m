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
%   `numbering_definitions` -- P8-1 FEATURE STUB: its faithful body builds a
%   _NumberingDefinitions over self._element, whose __len__ reads
%   self._numbering.num_lst -- requiring the CT_Numbering `num_lst` descriptor
%   (P8-1). Never on the open/save path.
%
%   ARG ORDER (docx): NumberingPart(partname, content_type, element, package).
%
%   Ported from python-docx v1.2.0: src/docx/parts/numbering.py::NumberingPart
%   (new 11-14 faithful NotImplementedError; numbering_definitions 16-20 -> P8-1
%   stub. The _NumberingDefinitions collection 23-32 lands at P8-1.)

    methods
        function obj = NumberingPart(partname, content_type, element, package)
            % Pass-through to the XmlPart constructor (design.md CT_*/part
            %   constructor contract): forward ALL args, no re-validation.
            obj@mat2doc.opc.XmlPart(partname, content_type, element, package);
        end

        function nd = numbering_definitions(obj) %#ok<MANU,STOUT>
            % NUMBERING_DEFINITIONS STUB (numbering.py 16-20, @lazyproperty).
            %   Owner: P8-1. Faithful body: return _NumberingDefinitions(self._element).
            error("mat2doc:notYetPorted", "%s", ...
                "mat2doc.parts._NumberingDefinitions / CT_Numbering.num_lst " + ...
                "(owning WP: P8-1 numbering tier) required by " + ...
                "mat2doc.parts.NumberingPart.numbering_definitions");
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
