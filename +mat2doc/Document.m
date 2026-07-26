function document = Document(docx)
% DOCUMENT Open or create a Word document.
%
%   d = MAT2DOC.DOCUMENT() opens the built-in default document "template" (the
%   bundled +mat2doc\templates\default.docx), mirroring python-docx's
%   default-template behavior.
%
%   d = MAT2DOC.DOCUMENT(docx) opens the .docx file at `docx` (a path string).
%   If the file's main document part is not a Word document part, a
%   mat2doc:ValueError is raised.
%
%   Inputs:
%       docx - (optional) path to a .docx file (string). Missing or the None
%              sentinel ([]) loads the built-in default template.
%   Outputs:
%       d    - a mat2doc.document.Document object.
%
%   Example:
%       d = mat2doc.Document();          % open the default template
%       d.save("out.docx");
%       d = mat2doc.Document("report.docx");
%
%   NOTE (module-vs-function): this file `+mat2doc\Document.m` is the public
%   entry function mat2doc.Document(...). The document PROXY class lives at
%   mat2doc.document.Document (`+mat2doc\+document\Document.m`); the two
%   fully-qualified names do not collide (design.md section 1 mirrors Python
%   docx.api.Document vs docx.document.Document).
%
%   Ported from python-docx v1.2.0: src/docx/api.py::Document (lines 19-37)

    arguments
        docx = []       % None sentinel (H3): default -> built-in template
    end

    % api.py 26: docx = _default_docx_path() if docx is None else docx
    % None-idiom: INLINE None test (RATIFIED at the P1->P2 boundary,
    % decision_2026-07-26_mat2doc_none_idiom.md; NO shared isNone helper).
    % STRICT form at this public entry (H3, Gate-2 fix): the sentinel is the
    % literal [] (0x0 double) ONLY. A bare isequal(docx, []) is true for the
    % 0x0 char '' as well, which would silently open the default template
    % where Python (docx='' is not None) attempts to open the path '' and
    % raises -- the exact "'' is NOT None" tri-state hazard. Mat2Ppt's
    % strict util.isNone at the identical api entry is the precedent.
    if strcmp(class(docx), 'double') && isequal(size(docx), [0, 0]) %#ok<STISA>
        docx = default_docx_path_();
    end

    % api.py 27: document_part = Package.open(docx).main_document_part
    % Opened through the docx-level Package subclass (docx.api imports
    % docx.package.Package), so every part receives a mat2doc.package.Package
    % back-reference and document_part.core_properties -> package.core_properties
    % resolves. The Python one-liner is split into two statements (MATLAB does
    % not chain a method call on a static-call result); behavior is identical.
    pkg = mat2doc.package.Package.open(docx);
    document_part = pkg.main_document_part();

    % api.py 28-30: reject a non-Word main document part.
    if document_part.content_type() ~= mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN
        error("mat2doc:ValueError", ...
            "file '%s' is not a Word file, content type is '%s'", ...
            docx, document_part.content_type());
    end

    % api.py 31: return document_part.document
    document = document_part.document();
end

function p = default_docx_path_()
% _default_docx_path (api.py 34-37): the path to the built-in default .docx
%   package, alongside this file under templates\default.docx. Resolved from the
%   package location (mfilename), NOT pwd, so mat2doc.Document() works from any
%   working directory (Python os.path.split(__file__)[0] analogue).
    thisdir = fileparts(mfilename('fullpath'));   % the +mat2doc directory
    p = string(fullfile(thisdir, "templates", "default.docx"));
end
