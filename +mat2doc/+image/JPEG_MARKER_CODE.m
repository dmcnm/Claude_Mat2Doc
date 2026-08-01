classdef JPEG_MARKER_CODE
% JPEG_MARKER_CODE JPEG marker codes and marker groupings.
%
%   Each single-byte marker code is held as a uint8 scalar (Python holds them as
%   length-1 bytes objects, e.g. b"\xe0"). The marker code that flows through the
%   JPEG parser is a 1x1 uint8 read off the stream (BytesIO.read(1)), so equality
%   compares directly against these uint8 constants.
%
%   STANDALONE_MARKERS and SOF_MARKER_CODES are uint8 vectors (Python tuples);
%   is_standalone / the SOF membership test use any(code == vector). marker_names
%   is a debug-only lookup (Python `marker_names` dict), used by Marker_.name
%   (# pragma: no cover upstream).
%
%   Ported from python-docx v1.2.0: src/docx/image/constants.py::JPEG_MARKER_CODE
%   (lines 4-97)

    properties (Constant)
        TEM  = uint8(1)     % b"\x01"
        DHT  = uint8(196)   % b"\xc4"
        DAC  = uint8(204)   % b"\xcc"
        JPG  = uint8(200)   % b"\xc8"

        SOF0 = uint8(192)   % b"\xc0"
        SOF1 = uint8(193)   % b"\xc1"
        SOF2 = uint8(194)   % b"\xc2"
        SOF3 = uint8(195)   % b"\xc3"
        SOF5 = uint8(197)   % b"\xc5"
        SOF6 = uint8(198)   % b"\xc6"
        SOF7 = uint8(199)   % b"\xc7"
        SOF9 = uint8(201)   % b"\xc9"
        SOFA = uint8(202)   % b"\xca"
        SOFB = uint8(203)   % b"\xcb"
        SOFD = uint8(205)   % b"\xcd"
        SOFE = uint8(206)   % b"\xce"
        SOFF = uint8(207)   % b"\xcf"

        RST0 = uint8(208)   % b"\xd0"
        RST1 = uint8(209)   % b"\xd1"
        RST2 = uint8(210)   % b"\xd2"
        RST3 = uint8(211)   % b"\xd3"
        RST4 = uint8(212)   % b"\xd4"
        RST5 = uint8(213)   % b"\xd5"
        RST6 = uint8(214)   % b"\xd6"
        RST7 = uint8(215)   % b"\xd7"

        SOI  = uint8(216)   % b"\xd8"
        EOI  = uint8(217)   % b"\xd9"
        SOS  = uint8(218)   % b"\xda"
        DQT  = uint8(219)   % b"\xdb"
        DNL  = uint8(220)   % b"\xdc"
        DRI  = uint8(221)   % b"\xdd"
        DHP  = uint8(222)   % b"\xde"
        EXP  = uint8(223)   % b"\xdf"

        APP0 = uint8(224)   % b"\xe0"
        APP1 = uint8(225)   % b"\xe1"
        APP2 = uint8(226)   % b"\xe2"
        APP3 = uint8(227)   % b"\xe3"
        APP4 = uint8(228)   % b"\xe4"
        APP5 = uint8(229)   % b"\xe5"
        APP6 = uint8(230)   % b"\xe6"
        APP7 = uint8(231)   % b"\xe7"
        APP8 = uint8(232)   % b"\xe8"
        APP9 = uint8(233)   % b"\xe9"
        APPA = uint8(234)   % b"\xea"
        APPB = uint8(235)   % b"\xeb"
        APPC = uint8(236)   % b"\xec"
        APPD = uint8(237)   % b"\xed"
        APPE = uint8(238)   % b"\xee"
        APPF = uint8(239)   % b"\xef"

        % STANDALONE_MARKERS = (TEM, SOI, EOI, RST0..RST7)  (constants.py 61)
        STANDALONE_MARKERS = uint8([1, 216, 217, 208, 209, 210, 211, 212, 213, 214, 215]);

        % SOF_MARKER_CODES = (SOF0,1,2,3,5,6,7,9,A,B,D,E,F)  (constants.py 63-77)
        SOF_MARKER_CODES = uint8([192, 193, 194, 195, 197, 198, 199, 201, 202, 203, 205, 206, 207]);
    end

    methods (Static)
        function tf = is_standalone(marker_code)
            % is_standalone (constants.py 95-97): marker_code in STANDALONE_MARKERS.
            tf = any(marker_code == mat2doc.image.JPEG_MARKER_CODE.STANDALONE_MARKERS);
        end

        function nm = marker_name(marker_code)
            % marker_names lookup (constants.py 79-93). Debug-only (Marker_.name is
            %   # pragma: no cover upstream). Raises KeyError for an unmapped code,
            %   matching Python dict indexing.
            switch marker_code
                case uint8(0),   nm = "UNKNOWN";
                case uint8(192), nm = "SOF0";
                case uint8(194), nm = "SOF2";
                case uint8(196), nm = "DHT";
                case uint8(218), nm = "SOS";
                case uint8(216), nm = "SOI";
                case uint8(217), nm = "EOI";
                case uint8(219), nm = "DQT";
                case uint8(224), nm = "APP0";
                case uint8(225), nm = "APP1";
                case uint8(226), nm = "APP2";
                case uint8(237), nm = "APP13";
                case uint8(238), nm = "APP14";
                otherwise
                    error("mat2doc:KeyError", "%s", ...
                        string(double(marker_code)));
            end
        end
    end
end
