function [tagSeq, S] = pprTagSeq()
% PPRTAGSEQ  Test-only CT_PPr._tag_seq successor-slice supplier (NOT toolbox code).
%
%   Supplies the child-descriptor ENGINE (the 11 generic methods on
%   mat2doc.oxml.BaseOxmlElement) with the real CT_PPr child ordering so the
%   H11 out-of-order build battery can be exercised WITHOUT a P4 CT_PPr class
%   (the CT_* classes do not exist yet). Mirrors the Gate-3 port driver
%   harness\mat2doc\validate_p1_3b.m verbatim.
%
%   Provenance: docx.oxml.text.parfmt.CT_PPr._tag_seq (python-docx v1.2.0,
%   parfmt.py 64-119) -- the 36-tag ordered child sequence. Python _tag_seq[N:]
%   maps to MATLAB tagSeq(N+1:end): the H1 (1-based) +1 base shift is applied
%   ONCE, here at declaration, so each successor slice below is already in
%   MATLAB indexing (see design.md H1).
%
%   Returns:
%     tagSeq  (1,36) string  -- the full ordered child tag list (w:*-prefixed)
%     S       struct         -- successor slices keyed by the 12 tags the H11
%                              battery drives; S.<tag> = _tag_seq[k+1:] i.e. the
%                              tags that must sort AFTER <tag>, in order.
%
%   Lives under tests\oxml\+p13btest (resolved by the Test class's TestClassSetup
%   PathFixture on tests\oxml); it is a test fixture, never Mat2Doc toolbox code
%   and never on the toolbox path.

    tagSeq = ["w:pStyle","w:keepNext","w:keepLines","w:pageBreakBefore","w:framePr", ...
        "w:widowControl","w:numPr","w:suppressLineNumbers","w:pBdr","w:shd","w:tabs", ...
        "w:suppressAutoHyphens","w:kinsoku","w:wordWrap","w:overflowPunct", ...
        "w:topLinePunct","w:autoSpaceDE","w:autoSpaceDN","w:bidi","w:adjustRightInd", ...
        "w:snapToGrid","w:spacing","w:ind","w:contextualSpacing","w:mirrorIndents", ...
        "w:suppressOverlap","w:jc","w:textDirection","w:textAlignment", ...
        "w:textboxTightWrap","w:outlineLvl","w:divId","w:cnfStyle","w:rPr", ...
        "w:sectPr","w:pPrChange"];
    assert(numel(tagSeq) == 36, "p13btest.pprTagSeq: tagSeq length must be 36");

    % Python _tag_seq[N:]  ->  MATLAB tagSeq(N+1:end)  (H1 base shift applied ONCE).
    S = struct( ...
        'pStyle',          tagSeq(2:end),  ...  % _tag_seq[1:]
        'keepNext',        tagSeq(3:end),  ...  % _tag_seq[2:]
        'keepLines',       tagSeq(4:end),  ...  % _tag_seq[3:]
        'pageBreakBefore', tagSeq(5:end),  ...  % _tag_seq[4:]
        'widowControl',    tagSeq(7:end),  ...  % _tag_seq[6:]
        'numPr',           tagSeq(8:end),  ...  % _tag_seq[7:]
        'tabs',            tagSeq(12:end), ...  % _tag_seq[11:]
        'spacing',         tagSeq(23:end), ...  % _tag_seq[22:]
        'ind',             tagSeq(24:end), ...  % _tag_seq[23:]
        'jc',              tagSeq(28:end), ...  % _tag_seq[27:]
        'outlineLvl',      tagSeq(32:end), ...  % _tag_seq[31:]
        'sectPr',          tagSeq(36:end));     % _tag_seq[35:]
end
