classdef CoreProperties < handle
% COREPROPERTIES Dublin-Core document properties for the package.
%
%   Corresponds to the part named /docProps/core.xml. Provides broadly-
%   standardized document metadata (author, last-modified, etc.) as read/write
%   properties, each delegating to the wrapped CT_CoreProperties element's
%   *_text / *_datetime / revision_number accessor (opc/coreprops.py 24-142).
%
%   HANDLE (design.md section 2): a proxy that wraps a shared element tree, so
%   two CoreProperties over the same element are views of one object -- Python
%   reference semantics.
%
%   The read/write accessors delegate entirely to the wrapped CT_CoreProperties
%   element, so its DEVIATIONS apply transitively: `revision` reads through
%   parse_int_ and `created`/`modified`/`last_printed` through the W3CDTF grammar
%   (both D-002, ASCII-only), and the created/modified setters carry the xsi
%   nsdecl hoist (D-serializer-nsdecl). See the CT_CoreProperties header +
%   validation\summary\deviation_ledger.md.
%
%   Example:
%       e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
%       cp = mat2doc.opc.CoreProperties(e);
%       cp.author = "Ada";                % <dc:creator>Ada</dc:creator>
%       disp(cp.author)                   % "Ada"
%       disp(cp.revision)                 % 0
%
%   Ported from python-docx v1.2.0: src/docx/opc/coreprops.py::CoreProperties

    properties (Access = private)
        element_    % the wrapped mat2doc.oxml.coreprops.CT_CoreProperties (opc/coreprops.py 21-22)
    end

    properties (Dependent)  % Dublin-Core read/write accessors (opc/coreprops.py 24-142)
        author
        category
        comments
        content_status
        created
        identifier
        keywords
        language
        last_modified_by
        last_printed
        modified
        revision
        subject
        title
        version
    end

    methods
        function obj = CoreProperties(element)
            % COREPROPERTIES Wrap a CT_CoreProperties element (opc/coreprops.py 21-22).
            obj.element_ = element;
        end

        % ---- author (creator) ---- opc/coreprops.py 24-30
        function v = get.author(obj);              v = obj.element_.author_text; end
        function set.author(obj, value);           obj.element_.author_text = value; end
        % ---- category ---- opc/coreprops.py 32-38
        function v = get.category(obj);            v = obj.element_.category_text; end
        function set.category(obj, value);         obj.element_.category_text = value; end
        % ---- comments (description) ---- opc/coreprops.py 40-46
        function v = get.comments(obj);            v = obj.element_.comments_text; end
        function set.comments(obj, value);         obj.element_.comments_text = value; end
        % ---- content_status (contentStatus) ---- opc/coreprops.py 48-54
        function v = get.content_status(obj);      v = obj.element_.contentStatus_text; end
        function set.content_status(obj, value);   obj.element_.contentStatus_text = value; end
        % ---- created ---- opc/coreprops.py 56-62
        function v = get.created(obj);             v = obj.element_.created_datetime; end
        function set.created(obj, value);          obj.element_.created_datetime = value; end
        % ---- identifier ---- opc/coreprops.py 64-70
        function v = get.identifier(obj);          v = obj.element_.identifier_text; end
        function set.identifier(obj, value);       obj.element_.identifier_text = value; end
        % ---- keywords ---- opc/coreprops.py 72-78
        function v = get.keywords(obj);            v = obj.element_.keywords_text; end
        function set.keywords(obj, value);         obj.element_.keywords_text = value; end
        % ---- language ---- opc/coreprops.py 80-86
        function v = get.language(obj);            v = obj.element_.language_text; end
        function set.language(obj, value);         obj.element_.language_text = value; end
        % ---- last_modified_by (lastModifiedBy) ---- opc/coreprops.py 88-94
        function v = get.last_modified_by(obj);    v = obj.element_.lastModifiedBy_text; end
        function set.last_modified_by(obj, value); obj.element_.lastModifiedBy_text = value; end
        % ---- last_printed (lastPrinted) ---- opc/coreprops.py 96-102
        function v = get.last_printed(obj);        v = obj.element_.lastPrinted_datetime; end
        function set.last_printed(obj, value);     obj.element_.lastPrinted_datetime = value; end
        % ---- modified ---- opc/coreprops.py 104-110
        function v = get.modified(obj);            v = obj.element_.modified_datetime; end
        function set.modified(obj, value);         obj.element_.modified_datetime = value; end
        % ---- revision ---- opc/coreprops.py 112-118
        function v = get.revision(obj);            v = obj.element_.revision_number; end
        function set.revision(obj, value);         obj.element_.revision_number = value; end
        % ---- subject ---- opc/coreprops.py 120-126
        function v = get.subject(obj);             v = obj.element_.subject_text; end
        function set.subject(obj, value);          obj.element_.subject_text = value; end
        % ---- title ---- opc/coreprops.py 128-134
        function v = get.title(obj);               v = obj.element_.title_text; end
        function set.title(obj, value);            obj.element_.title_text = value; end
        % ---- version ---- opc/coreprops.py 136-142
        function v = get.version(obj);             v = obj.element_.version_text; end
        function set.version(obj, value);          obj.element_.version_text = value; end
    end
end
