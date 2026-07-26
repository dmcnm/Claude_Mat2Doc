classdef XsdStringEnumeration < mat2doc.oxml.simpletypes.BaseStringEnumerationType
% XSDSTRINGENUMERATION Set of enumerated xsd:string values.
%
%   Abstract-style base (Python `class XsdStringEnumeration(...)` with only a
%   docstring): concrete subclasses (ST_HexColorAuto, ST_Merge,
%   ST_VerticalAlignRun) supply the member set and their own validate/to_xml.
%   Never used directly (it has no members), matching Python where calling the
%   base validate would fail on the missing _members.
%
%   Ported from python-docx v1.2.0:
%   src/docx/oxml/simpletypes.py::XsdStringEnumeration (lines 158-159)
end
