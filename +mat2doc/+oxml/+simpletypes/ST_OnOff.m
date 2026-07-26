classdef ST_OnOff < mat2doc.oxml.simpletypes.XsdBoolean
% ST_ONOFF The OOXML on/off boolean: '1'/'0'/'true'/'false'/'on'/'off'.
%
%   Widens XsdBoolean's accepted XML token set to also allow 'on'/'off'.
%   convert_from_xml raises the docx.exceptions.InvalidXmlError (via the
%   canonical raiser mat2doc.exc.InvalidXmlError) on an unrecognized token,
%   and returns True for '1'/'true'/'on'. Overrides ONLY convert_from_xml, so
%   re-declares from_xml to route to THIS convert_from_xml (H10);
%   validate / convert_to_xml / to_xml are inherited from XsdBoolean unchanged
%   (True/False validation; True->"1", False->"0").
%
%   Example:
%       disp(mat2doc.oxml.simpletypes.ST_OnOff.from_xml("on"))    % 1 (logical)
%       disp(mat2doc.oxml.simpletypes.ST_OnOff.from_xml("off"))   % 0 (logical)
%       disp(mat2doc.oxml.simpletypes.ST_OnOff.to_xml(true))      % "1"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_OnOff
%   (lines 336-344)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27), bound to THIS
            %   convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_OnOff.convert_from_xml(xml_value);
        end

        function v = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 337-344: one of '1','0','true','false',
            %   'on','off', else InvalidXmlError; True for '1'/'true'/'on'.
            %   Message verbatim (the source splits the literal across two lines
            %   but concatenates to a single string).
            sv = string(str_value);
            if ~any(sv == ["1", "0", "true", "false", "on", "off"])
                mat2doc.exc.InvalidXmlError( ...
                    "value must be one of '1', '0', 'true', 'false', 'on', or 'off', got '" ...
                    + sv + "'");
            end
            v = any(sv == ["1", "true", "on"]);
        end
    end
end
