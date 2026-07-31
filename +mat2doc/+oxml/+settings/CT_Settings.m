classdef CT_Settings < mat2doc.oxml.BaseOxmlElement
% CT_SETTINGS `<w:settings>` element, root element for the settings part.
%
%   Carries a single ported ZeroOrOne child descriptor -- evenAndOddHeaders (a
%   CT_OnOff) -- plus the derived boolean @property evenAndOddHeaders_val
%   (get/set). Registered for w:settings (docx/oxml/__init__.py 132-134); the
%   settings part (word/settings.xml) transits this class on every load.
%
%   H11 SUCCESSOR-SLICE (THE correctness crux). _tag_seq (settings.py 19-118,
%   VERBATIM, 98 tags) is stored as the Constant TAG_SEQ. The descriptor uses
%   Python successors=_tag_seq[48:]; the own tag `w:evenAndOddHeaders` sits at
%   Python 0-based index 47 = MATLAB 1-based index 48 = TAG_SEQ(48). H1 slice
%   map: Python _tag_seq[48:] -> MATLAB TAG_SEQ(49:end) (0-based slice start 48
%   -> 1-based start 49). TAG_SEQ(49) == "w:bookFoldRevPrinting" (the tag
%   immediately AFTER evenAndOddHeaders), so insert_element_before re-sorts a
%   scrambled evenAndOddHeaders add into canonical OOXML schema order (right
%   after w:defaultTableStyle, before w:bookFoldRevPrinting). A wrong slice here
%   inserts it in the wrong position -> Word repair / byte divergence.
%
%   The m:mathPr (index 84, 1-based) and sl:schemaLibrary (index 94, 1-based)
%   prefixes ARE present in +oxml/nsmap.m (m/sl), so successor tags carrying
%   those prefixes resolve correctly (H8).
%
%   GENERATED DESCRIPTOR FAMILY (ZeroOrOne, xmlchemy docx form): get.x,
%   get_or_add_x, new_x_, insert_x_, add_x_, remove_x_ (underscore rotation:
%   Python _new_x/_insert_x/_add_x/_remove_x -> new_x_/insert_x_/add_x_/
%   remove_x_; get_or_add_x public). The pyright Callable annotations
%   (settings.py 16-17: get_or_add_evenAndOddHeaders / _remove_evenAndOddHeaders)
%   are type hints only and add no members beyond this family. The descriptor
%   uses the generic BaseOxmlElement engine (no _new_/_insert_ override).
%
%   CHILD-CLASS REGISTRATION: w:evenAndOddHeaders -> CT_OnOff
%   (docx/oxml/__init__.py:83; registered by P5-1). evenAndOddHeaders_val
%   reads/writes .val on that CT_OnOff child.
%
%   H3 (tri-state) / D-delta-1: evenAndOddHeaders_val is derived from the
%   CT_OnOff-backed presence of w:evenAndOddHeaders (None/True/False). Getter:
%   child absent -> False (settings.py 128-129, NOT the child's own default).
%   Setter (settings.py 132-138): value is None OR value is False -> REMOVE the
%   child; else get_or_add + set .val = value (True -> CT_OnOff.val setter
%   removes @val, emitting the empty <w:evenAndOddHeaders/>, per D-delta-1).
%
%   H4 (truthiness): Python `value is None or value is False` is IDENTITY, not
%   truthiness -- a double 0 must NOT match `is False`. Ported with the house
%   islogical-guard idiom (Font.m 315-317): isequal(value,[]) OR (islogical &&
%   isscalar && ~value).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the <w:settings> root of word/settings.xml, so all
%   positional args forward verbatim.
%
%   Example:
%       s = mat2doc.oxml.OxmlElement("w:settings");   % a CT_Settings
%       s.evenAndOddHeaders_val                         % false (child absent)
%       s.evenAndOddHeaders_val = true;                 % -> <w:evenAndOddHeaders/>
%       s.evenAndOddHeaders_val = false;                % removes the child
%
%   Ported from python-docx v1.2.0: src/docx/oxml/settings.py::CT_Settings
%   (lines 13-138; registered for w:settings)

    properties (Constant, Hidden)  % _tag_seq VERBATIM (settings.py 19-118; 98 tags)
        TAG_SEQ = [ ...
            "w:writeProtection", "w:view", "w:zoom", "w:removePersonalInformation", ...      %  1- 4
            "w:removeDateAndTime", "w:doNotDisplayPageBoundaries", ...                        %  5- 6
            "w:displayBackgroundShape", "w:printPostScriptOverText", ...                      %  7- 8
            "w:printFractionalCharacterWidth", "w:printFormsData", ...                        %  9-10
            "w:embedTrueTypeFonts", "w:embedSystemFonts", "w:saveSubsetFonts", ...            % 11-13
            "w:saveFormsData", "w:mirrorMargins", "w:alignBordersAndEdges", ...               % 14-16
            "w:bordersDoNotSurroundHeader", "w:bordersDoNotSurroundFooter", ...               % 17-18
            "w:gutterAtTop", "w:hideSpellingErrors", "w:hideGrammaticalErrors", ...           % 19-21
            "w:activeWritingStyle", "w:proofState", "w:formsDesign", ...                      % 22-24
            "w:attachedTemplate", "w:linkStyles", "w:stylePaneFormatFilter", ...              % 25-27
            "w:stylePaneSortMethod", "w:documentType", "w:mailMerge", ...                     % 28-30
            "w:revisionView", "w:trackRevisions", "w:doNotTrackMoves", ...                    % 31-33
            "w:doNotTrackFormatting", "w:documentProtection", "w:autoFormatOverride", ...     % 34-36
            "w:styleLockTheme", "w:styleLockQFSet", "w:defaultTabStop", ...                   % 37-39
            "w:autoHyphenation", "w:consecutiveHyphenLimit", "w:hyphenationZone", ...         % 40-42
            "w:doNotHyphenateCaps", "w:showEnvelope", "w:summaryLength", ...                  % 43-45
            "w:clickAndTypeStyle", "w:defaultTableStyle", "w:evenAndOddHeaders", ...          % 46-48  (own tag @48)
            "w:bookFoldRevPrinting", "w:bookFoldPrinting", "w:bookFoldPrintingSheets", ...    % 49-51  (successors start @49)
            "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing", ...            % 52-53
            "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery", ...   % 54-55
            "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin", ...     % 56-57
            "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", ...                        % 58-59
            "w:noPunctuationKerning", "w:characterSpacingControl", "w:printTwoOnOne", ...     % 60-62
            "w:strictFirstAndLastChars", "w:noLineBreaksAfter", "w:noLineBreaksBefore", ...   % 63-65
            "w:savePreviewPicture", "w:doNotValidateAgainstSchema", "w:saveInvalidXml", ...   % 66-68
            "w:ignoreMixedContent", "w:alwaysShowPlaceholderText", ...                        % 69-70
            "w:doNotDemarcateInvalidXml", "w:saveXmlDataOnly", "w:useXSLTWhenSaving", ...     % 71-73
            "w:saveThroughXslt", "w:showXMLTags", "w:alwaysMergeEmptyNamespace", ...          % 74-76
            "w:updateFields", "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", ...        % 77-80
            "w:compat", "w:docVars", "w:rsids", "m:mathPr", "w:attachedSchema", ...           % 81-85
            "w:themeFontLang", "w:clrSchemeMapping", "w:doNotIncludeSubdocsInStats", ...      % 86-88
            "w:doNotAutoCompressPictures", "w:forceUpgrade", "w:captions", ...                % 89-91
            "w:readModeInkLockDown", "w:smartTagType", "sl:schemaLibrary", ...                % 92-94
            "w:shapeDefaults", "w:doNotEmbedSmartTags", "w:decimalSymbol", ...                % 95-97
            "w:listSeparator" ]                                                               % 98
    end

    properties (Dependent)  % generated ZeroOrOne getter + @property member
        evenAndOddHeaders       % ZeroOrOne child (read-only; use get_or_add/remove)
        evenAndOddHeaders_val   % bool from w:evenAndOddHeaders presence (False if absent)
    end

    methods
        function obj = CT_Settings(varargin)
            % CT_SETTINGS Construct a loose <w:settings> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- evenAndOddHeaders (ZeroOrOne, successors=_tag_seq[48:] -> TAG_SEQ(49:end)) ----
        function child = get.evenAndOddHeaders(obj);            child = obj.getChild("w:evenAndOddHeaders"); end
        function child = get_or_add_evenAndOddHeaders(obj);     child = obj.getOrAddChild("w:evenAndOddHeaders", obj.TAG_SEQ(49:end)); end
        function child = new_evenAndOddHeaders_(obj);           child = obj.newChild("w:evenAndOddHeaders"); end
        function child = insert_evenAndOddHeaders_(obj, child); child = obj.insertChildInSequence(child, obj.TAG_SEQ(49:end)); end
        function child = add_evenAndOddHeaders_(obj, varargin); child = obj.addChild("w:evenAndOddHeaders", obj.TAG_SEQ(49:end), varargin{:}); end
        function remove_evenAndOddHeaders_(obj);                obj.removeChild("w:evenAndOddHeaders"); end

        % ---- evenAndOddHeaders_val (@property, settings.py 124-138) ----
        function value = get.evenAndOddHeaders_val(obj)
            % Python (settings.py 127-130):
            %   evenAndOddHeaders = self.evenAndOddHeaders
            %   if evenAndOddHeaders is None: return False
            %   return evenAndOddHeaders.val
            e = obj.evenAndOddHeaders;
            if isequal(e, [])          % Python: if evenAndOddHeaders is None
                value = false;
                return
            end
            value = e.val;
        end
        function set.evenAndOddHeaders_val(obj, value)
            % Python (settings.py 133-138):
            %   if value is None or value is False:
            %       self._remove_evenAndOddHeaders(); return
            %   self.get_or_add_evenAndOddHeaders().val = value
            % H4: `is None or is False` is IDENTITY -- a double 0 must NOT match
            % (house islogical-guard idiom, Font.m 315-317).
            if isequal(value, []) || (islogical(value) && isscalar(value) && ~value)
                obj.remove_evenAndOddHeaders_();
                return
            end
            e = obj.get_or_add_evenAndOddHeaders();
            e.val = value;
        end
    end
end
