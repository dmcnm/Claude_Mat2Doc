classdef ST_DateTime < mat2doc.oxml.simpletypes.BaseSimpleType
% ST_DATETIME xsd:dateTime for `w:` attributes (e.g. w:comment/@w:date).
%
%   DOCX-NOVEL validator (no python-pptx analogue). Represents a Python
%   datetime.datetime as a MATLAB `datetime`. A tz-aware Python datetime maps
%   to a MATLAB datetime with a TimeZone (always normalized to 'UTC' here,
%   see VERIFY-tz); a NAIVE Python datetime (tzinfo is None) maps to a MATLAB
%   datetime with an EMPTY TimeZone. The Z-suffixed and offset forms are
%   tz-aware; the plain "yyyy-mm-ddThh:mm:ss" and date-only forms are naive.
%   Extends BaseSimpleType directly, so it declares its own from_xml / to_xml
%   (H10). Used by comments.py (w:comment/@w:date).
%
%   Distinct from the coreprops W3CDTF handling (CT_CoreProperties parses/
%   formats cp:created etc. itself): this is xsd:dateTime with a different
%   grammar and a NEVER-RAISE parse (any parse failure yields the 1970 epoch).
%
%   PARSE (convert_from_xml, simpletypes.py 219-249) -- Python parse_xsd_datetime:
%     1. If the string ends with 'Z': try "%Y-%m-%dT%H:%M:%S.%fZ" (fractional),
%        else "%Y-%m-%dT%H:%M:%SZ" -> tz-aware UTC. (Terminal: no fall-through.)
%     2. Else dt.datetime.fromisoformat(str) -- the broad ISO-8601 parser.
%     3. Else "%Y-%m-%dT%H:%M:%S.%f" then "%Y-%m-%dT%H:%M:%S" -> NAIVE.
%     Any exception from the whole chain -> dt.datetime(1970,1,1, tz=UTC).
%
%   VERIFY-fromiso (raised for the Auditor/Validator): Python step 2 is
%   dt.datetime.fromisoformat, which in CPython 3.11+ accepts a VERY broad ISO
%   surface (basic no-separator format "YYYYMMDD", any single-char date/time
%   separator, time without seconds "HH:MM" or hour-only "HH", week dates,
%   ordinal dates, offset seconds/fractions, trailing 'Z'). MATLAB has no
%   fromisoformat. fromIsoFormat_ below implements the EXTENDED-format subset
%   that real Word `w:date` values use: date-only "YYYY-MM-DD", and
%   "YYYY-MM-DD[T ]HH:MM:SS[.ffffff][(+/-)HH:MM]". Inputs OUTSIDE that subset
%   (seconds omitted, basic format, week/ordinal dates, offset seconds) that
%   Python fromisoformat WOULD parse fall through to the naive templates and,
%   failing those, to the 1970 epoch -- a DIVERGENCE from Python for those
%   exotic inputs (Python returns the parsed datetime, the port returns epoch).
%   Direction is a safe under-accept for realistic w:date; flagged for the
%   Auditor to re-verify against the oracle and, if any real corpus value hits
%   it, to widen the subset or open a fresh D-number.
%
%   VERIFY-tz: a tz-aware input with a non-UTC offset (e.g. "...-08:00") is
%   stored as the equivalent UTC instant (TimeZone 'UTC'), not with the
%   original offset. Python keeps the original fixed-offset tzinfo. Because
%   convert_to_xml normalizes to UTC before formatting (astimezone(utc)), the
%   emitted XML bytes are IDENTICAL; only the intermediate API datetime VALUE's
%   zone label differs (same instant). Same class of tz API-value divergence
%   the coreprops port documented; no XML-byte effect.
%
%   D-002: the date grammar accepts ASCII digits [0-9] only, where CPython
%   strptime/fromisoformat also accept Unicode decimal digits. Safe
%   under-accept; adopted D-002 (no new D-number).
%
%   FORMAT (convert_to_xml, simpletypes.py 251-261): a naive value is first
%   interpreted as LOCAL time (astimezone() with no arg -- ENVIRONMENT
%   dependent, faithful to Python), then converted to UTC, then formatted
%   "%Y-%m-%dT%H:%M:%SZ" (whole seconds, literal Z).
%
%   Example:
%       t = mat2doc.oxml.simpletypes.ST_DateTime.from_xml("2023-01-02T03:04:05Z");
%       mat2doc.oxml.simpletypes.ST_DateTime.to_xml(t)   % "2023-01-02T03:04:05Z"
%
%   Ported from python-docx v1.2.0: src/docx/oxml/simpletypes.py::ST_DateTime
%   (lines 217-266)

    methods (Static)
        function v = from_xml(xml_value)
            % FROM_XML BaseSimpleType.from_xml (lines 25-27): convert_from_xml.
            v = mat2doc.oxml.simpletypes.ST_DateTime.convert_from_xml(xml_value);
        end

        function s = to_xml(value)
            % TO_XML BaseSimpleType.to_xml (lines 29-33): validate; convert_to_xml.
            mat2doc.oxml.simpletypes.ST_DateTime.validate(value);
            s = mat2doc.oxml.simpletypes.ST_DateTime.convert_to_xml(value);
        end

        function ts = convert_from_xml(str_value)
            % CONVERT_FROM_XML lines 219-249: parse, but NEVER raise -- any
            %   failure yields the 1970 UTC epoch (Python `except Exception`).
            try
                ts = mat2doc.oxml.simpletypes.ST_DateTime.parseXsdDatetime_(string(str_value));
            catch
                ts = datetime(1970, 1, 1, 0, 0, 0, "TimeZone", "UTC");  % dt.datetime(1970,1,1,tzinfo=utc)
            end
        end

        function s = convert_to_xml(value)
            % CONVERT_TO_XML lines 251-261: naive -> local -> UTC; strftime
            %   "%Y-%m-%dT%H:%M:%SZ" (whole seconds + literal Z).
            v = value;
            if isempty(v.TimeZone)          % value.tzinfo is None (naive)
                v.TimeZone = "local";       % astimezone() assumes system local
            end
            v.TimeZone = "UTC";             % astimezone(dt.timezone.utc)
            s = string(v, "yyyy-MM-dd'T'HH:mm:ss") + "Z";
        end

        function validate(value)
            % VALIDATE lines 263-266: must be a datetime, else TypeError
            %   "only a datetime.datetime object may be assigned, got '%s'"
            %   % value. Value rendered best-effort; class token per D-005.
            if ~isa(value, "datetime")
                error("mat2doc:TypeError", ...
                    "only a datetime.datetime object may be assigned, got '%s'", ...
                    mat2doc.oxml.simpletypes.ST_DateTime.valueRepr_(value));
            end
        end
    end

    methods (Static, Access = private)
        function ts = parseXsdDatetime_(s)
            % PARSEXSDDATETIME_ lines 222-243. Raises mat2doc:ValueError on
            %   failure (caught by convert_from_xml -> epoch).
            import mat2doc.oxml.simpletypes.ST_DateTime
            if endsWith(s, "Z")
                % -- trailing 'Z' (Zulu/UTC) branch (terminal) --
                tok = regexp(s, "^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{1,6})Z$", ...
                    "tokens", "once");
                if ~isempty(tok)
                    f = str2double(tok(1:6));
                    frac = str2double("0." + string(tok{7}));
                    ts = ST_DateTime.buildValidated_(f, frac);
                    ts.TimeZone = "UTC";
                    return
                end
                tok = regexp(s, "^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})Z$", ...
                    "tokens", "once");
                if ~isempty(tok)
                    f = str2double(tok(1:6));
                    ts = ST_DateTime.buildValidated_(f, 0);
                    ts.TimeZone = "UTC";
                    return
                end
                error("mat2doc:ValueError", "unparseable Z datetime");
            end
            % -- fromisoformat branch (explicit offsets / naive) --
            try
                ts = ST_DateTime.fromIsoFormat_(s);
                return
            catch
                % fall through to the naive strptime templates
            end
            % -- naive strptime fallback (with or without fractional seconds) --
            tok = regexp(s, "^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})\.(\d{1,6})$", ...
                "tokens", "once");
            if ~isempty(tok)
                f = str2double(tok(1:6));
                frac = str2double("0." + string(tok{7}));
                ts = ST_DateTime.buildValidated_(f, frac);   % NAIVE (no TimeZone)
                return
            end
            tok = regexp(s, "^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$", ...
                "tokens", "once");
            if ~isempty(tok)
                f = str2double(tok(1:6));
                ts = ST_DateTime.buildValidated_(f, 0);      % NAIVE
                return
            end
            error("mat2doc:ValueError", "unparseable naive datetime");
        end

        function ts = fromIsoFormat_(s)
            % FROMISOFORMAT_ Extended-format subset of dt.datetime.fromisoformat
            %   (see VERIFY-fromiso in the class header). Raises mat2doc:ValueError
            %   for anything outside the supported subset.
            import mat2doc.oxml.simpletypes.ST_DateTime
            % date-only -> naive midnight
            tok = regexp(s, "^(\d{4})-(\d{2})-(\d{2})$", "tokens", "once");
            if ~isempty(tok)
                f = [str2double(tok), 0, 0, 0];
                ts = ST_DateTime.buildValidated_(f, 0);      % NAIVE
                return
            end
            % datetime with optional fractional seconds and optional +/-HH:MM
            % offset. NOTE: flat capture groups only -- a NESTED optional group
            % ((\.(\d{1,6}))?) collapses to a single token in MATLAB regexp when
            % it does not participate, shifting later group indices. So group 7
            % captures the whole ".ffffff" (dot included), group 8 the offset.
            tok = regexp(s, ...
                "^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})(\.\d{1,6})?([+-]\d{2}:\d{2})?$", ...
                "tokens", "once");
            if isempty(tok)
                error("mat2doc:ValueError", "unsupported isoformat");
            end
            f = str2double(tok(1:6));
            fracStr = string(tok{7});           % ".ffffff" or ""
            if strlength(fracStr) > 0
                frac = str2double("0" + fracStr);   % ".123" -> "0.123"
            else
                frac = 0;
            end
            ts = ST_DateTime.buildValidated_(f, frac);
            off = string(tok{8});
            if strlength(off) > 0
                % Explicit offset -> tz-aware. Shift the local wall clock to the
                % equivalent UTC instant and label it UTC (VERIFY-tz). Offset
                % "+05:00" means local is 5h ahead of UTC, so UTC = local - 5h.
                sgn = 1;
                if extractBefore(off, 2) == "-"
                    sgn = -1;
                end
                offh = str2double(extractBetween(off, 2, 3));
                offm = str2double(extractBetween(off, 5, 6));
                ts = ts - sgn * (hours(offh) + minutes(offm));
                ts.TimeZone = "UTC";
            end
            % else: naive (no offset) -- leave TimeZone empty
        end

        function ts = buildValidated_(f, frac)
            % BUILDVALIDATED_ Construct a NAIVE datetime from integer fields
            %   f = [Y Mo D H Mi SS] plus a fractional-seconds part, validating
            %   ranges by component round-trip (a rolled-over field means a
            %   field was out of range -> strptime/datetime ValueError). Mirrors
            %   the coreprops parse_W3CDTF_to_datetime_ validation idiom (Mat2Ppt
            %   Gate-2 F2/F2b), incl. the CPython MINYEAR=1 floor.
            Y = f(1); Mo = f(2); D = f(3); H = f(4); Mi = f(5); SS = f(6);
            if Y < 1
                % CPython datetime MINYEAR = 1: year 0 raises. MATLAB datetime
                % accepts year 0 and would re-emit "0001-..."; reject like Python.
                error("mat2doc:ValueError", "year is out of range");
            end
            cand = datetime(Y, Mo, D, H, Mi, SS);
            [cy, cmo, cd] = ymd(cand);
            [ch, cmi, cs] = hms(cand);
            if ~isequal([cy cmo cd ch cmi cs], [Y Mo D H Mi SS])
                error("mat2doc:ValueError", "datetime field out of range");
            end
            if frac ~= 0
                cand = cand + seconds(frac);
            end
            ts = cand;   % NAIVE (TimeZone empty); caller labels/shifts as needed
        end

        function r = valueRepr_(value)
            % Best-effort str(value) for the validate TypeError message.
            try
                r = mat2doc.shared.pyStr(value);
            catch
                r = string(class(value));
            end
        end
    end
end
