classdef CT_Ind < mat2doc.oxml.BaseOxmlElement
% CT_IND `<w:ind>` element, specifying paragraph indentation.
%
%   Four OptionalAttributes, all default None ([]):
%     left      OptionalAttribute("w:left",      ST_SignedTwipsMeasure)  parfmt.py:32
%     right     OptionalAttribute("w:right",     ST_SignedTwipsMeasure)  parfmt.py:35
%     firstLine OptionalAttribute("w:firstLine", ST_TwipsMeasure)        parfmt.py:38
%     hanging   OptionalAttribute("w:hanging",   ST_TwipsMeasure)        parfmt.py:41
%
%   VERIFIED per attr (parfmt.py 32-43): left/right use ST_SignedTwipsMeasure
%   (signed), firstLine/hanging use ST_TwipsMeasure (unsigned). All -> Length.
%
%   H3 tri-state: no attribute has a Python default (default None -> []); the
%   getter returns [] when absent and the setter removes the attribute when
%   assigned [] (None) -- getAttrTyped/setAttrTyped handle this.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:ind> nodes inside a real part.
%
%   Example:
%       ind = mat2doc.oxml.OxmlElement("w:ind");
%       ind.left = mat2doc.shared.Twips(720);   % <w:ind w:left="720"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_Ind
%   (lines 29-43; registered for w:ind)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        LEFT_ATTR       = "w:left"                % OptionalAttribute @ parfmt.py:32
        LEFT_TYPE       = "ST_SignedTwipsMeasure"
        LEFT_DEFAULT    = []                       % Python default: None
        RIGHT_ATTR      = "w:right"               % OptionalAttribute @ parfmt.py:35
        RIGHT_TYPE      = "ST_SignedTwipsMeasure"
        RIGHT_DEFAULT   = []                       % Python default: None
        FIRSTLINE_ATTR  = "w:firstLine"           % OptionalAttribute @ parfmt.py:38
        FIRSTLINE_TYPE  = "ST_TwipsMeasure"
        FIRSTLINE_DEFAULT = []                     % Python default: None
        HANGING_ATTR    = "w:hanging"             % OptionalAttribute @ parfmt.py:41
        HANGING_TYPE    = "ST_TwipsMeasure"
        HANGING_DEFAULT = []                       % Python default: None
    end

    properties (Dependent)  % generated descriptor properties
        left       % OptionalAttribute('w:left', ST_SignedTwipsMeasure) -> Length or []
        right      % OptionalAttribute('w:right', ST_SignedTwipsMeasure) -> Length or []
        firstLine  % OptionalAttribute('w:firstLine', ST_TwipsMeasure) -> Length or []
        hanging    % OptionalAttribute('w:hanging', ST_TwipsMeasure) -> Length or []
    end

    methods
        function obj = CT_Ind(varargin)
            % CT_IND Construct a loose <w:ind> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- left (OptionalAttribute, ST_SignedTwipsMeasure, default None) ----
        function value = get.left(obj)
            value = obj.getAttrTyped(obj.LEFT_ATTR, obj.LEFT_TYPE, obj.LEFT_DEFAULT);
        end
        function set.left(obj, value)
            obj.setAttrTyped(obj.LEFT_ATTR, obj.LEFT_TYPE, value, obj.LEFT_DEFAULT);
        end

        % ---- right (OptionalAttribute, ST_SignedTwipsMeasure, default None) ----
        function value = get.right(obj)
            value = obj.getAttrTyped(obj.RIGHT_ATTR, obj.RIGHT_TYPE, obj.RIGHT_DEFAULT);
        end
        function set.right(obj, value)
            obj.setAttrTyped(obj.RIGHT_ATTR, obj.RIGHT_TYPE, value, obj.RIGHT_DEFAULT);
        end

        % ---- firstLine (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.firstLine(obj)
            value = obj.getAttrTyped(obj.FIRSTLINE_ATTR, obj.FIRSTLINE_TYPE, obj.FIRSTLINE_DEFAULT);
        end
        function set.firstLine(obj, value)
            obj.setAttrTyped(obj.FIRSTLINE_ATTR, obj.FIRSTLINE_TYPE, value, obj.FIRSTLINE_DEFAULT);
        end

        % ---- hanging (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.hanging(obj)
            value = obj.getAttrTyped(obj.HANGING_ATTR, obj.HANGING_TYPE, obj.HANGING_DEFAULT);
        end
        function set.hanging(obj, value)
            obj.setAttrTyped(obj.HANGING_ATTR, obj.HANGING_TYPE, value, obj.HANGING_DEFAULT);
        end
    end
end
