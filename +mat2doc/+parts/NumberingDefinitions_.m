classdef NumberingDefinitions_ < handle
% NUMBERINGDEFINITIONS_ Collection of the <w:num> numbering-definition proxies.
%
%   Corresponds to the <w:num> elements in a numbering part. In python-docx this
%   is `_NumberingDefinitions`, a plain object (NOT an ElementProxy) that wraps the
%   <w:numbering> element and reports how many <w:num> definitions it holds.
%   Underscore rotation: leading `_` -> trailing (`_NumberingDefinitions` ->
%   `NumberingDefinitions_`), matching LatentStyle_ / TableStyle_ precedent.
%   Modeled as a `handle` (Python reference semantics; two references view one
%   object).
%
%   v1.2.0 SURFACE: EXACTLY __init__ and __len__. There is NO __getitem__ /
%   __iter__ / add member -- so no RedefinesParen () access is needed. `len(nd)`
%   is ported as len_() (design.md collection idiom; LatentStyles precedent):
%   `len(x)` -> x.len_().
%
%   Example:
%       dp = pkg.main_document_part();
%       nd = dp.numbering_part.numbering_definitions;   % (numbering_part raises
%                                                       %  NotImplementedError until
%                                                       %  the part already exists)
%       nd.len_()                                       % number of <w:num> defs
%
%   Ported from python-docx v1.2.0: src/docx/parts/numbering.py::_NumberingDefinitions
%   (lines 23-32)

    properties (Access = private)
        numbering_   % the wrapped <w:numbering> element (a CT_Numbering)
    end

    methods
        function obj = NumberingDefinitions_(numbering_elm)
            % NUMBERINGDEFINITIONS_ Wrap a <w:numbering> element (numbering.py 27-29).
            %   Python: super().__init__(); self._numbering = numbering_elm
            %
            %   Inputs:  numbering_elm - a mat2doc.oxml.numbering.CT_Numbering.
            %   Outputs: obj           - a scalar NumberingDefinitions_ handle.
            obj.numbering_ = numbering_elm;
        end

        function n = len_(obj)
            % LEN_ Number of <w:num> definitions (numbering.py 31-32).
            %   Python __len__: return len(self._numbering.num_lst)
            n = numel(obj.numbering_.num_lst);
        end
    end
end
