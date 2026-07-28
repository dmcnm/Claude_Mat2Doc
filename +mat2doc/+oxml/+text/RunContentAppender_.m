classdef RunContentAppender_ < handle
% RUNCONTENTAPPENDER_ Translate a string into run-content elements in a <w:r>.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::_RunContentAppender
%   (lines 262-307). The module-private helper class `_RunContentAppender` rotates
%   the leading underscore to the END -> RunContentAppender_ (design.md section 2).
%
%   Contiguous sequences of regular characters are appended in a single <w:t>
%   element. Each tab character ("\t") appends a <w:tab/>; each newline or
%   carriage return ("\n"/"\r") appends a <w:br/> (via CT_R.add_br; the schema
%   uses w:br for the soft break produced here -- run.py 297-299). A finite-state
%   machine buffers regular characters until a \t/\r/\n boundary or end of text,
%   where flush() writes the pending <w:t>.
%
%   REFERENCE SEMANTICS: a handle class -- add_char/flush mutate the shared buffer
%   and the shared CT_R, mirroring the mutable Python instance.
%
%   UNDERSCORE ROTATION: the private `_r` / `_bfr` attributes rotate the leading
%   underscore -> r_ / bfr_.
%
%   H2 (UTF-8 / iteration): Python iterates `text` by Unicode code point. MATLAB
%   iterates the UTF-16 char units of char(text). A non-BMP (astral) character is
%   two UTF-16 surrogate units; since neither surrogate half is \t/\r/\n, both are
%   buffered and rejoined by string(bfr_) on flush -- byte-equivalent to Python's
%   single-code-point handling (the split is unobservable, it only matters at
%   \t/\r/\n boundaries, which are single BMP chars).
%
%   H4 (truthiness): flush's `if text:` (append only a NON-empty <w:t>) ->
%   `if strlength(text) > 0`.
%
%   Example:
%       r = mat2doc.oxml.OxmlElement("w:r");
%       mat2doc.oxml.text.RunContentAppender_.append_to_run_from_text(r, "a\tb");
%       % r now has <w:t>a</w:t><w:tab/><w:t>b</w:t>

    properties (Access = private)
        r_          % the CT_R (<w:r>) being appended to
        bfr_ char = ''   % _bfr: buffered regular characters (pending <w:t>)
    end

    methods
        function obj = RunContentAppender_(r)
            % RUNCONTENTAPPENDER_ Construct over a CT_R (run.py 271-273):
            %   self._r = r; self._bfr = [].
            obj.r_ = r;
            obj.bfr_ = '';
        end

        function add_text(obj, text)
            % ADD_TEXT Append inner-content elements for `text` (run.py 281-285).
            %   Python: for char in text: self.add_char(char); self.flush().
            arguments
                obj
                text (1,1) string
            end
            chars = char(text);              % UTF-16 code units (H2)
            for k = 1:numel(chars)
                obj.add_char(chars(k));
            end
            obj.flush();
        end

        function add_char(obj, ch)
            % ADD_CHAR Process the next input character through the FSM
            %   (run.py 287-301). \t -> flush + add_tab; \r|\n -> flush + add_br;
            %   else buffer the character.
            if ch == char(9)                       % Python: if char == "\t"
                obj.flush();
                obj.r_.add_tab();
            elseif ch == char(13) || ch == char(10) % Python: elif char in "\r\n"
                obj.flush();
                obj.r_.add_br();
            else
                obj.bfr_(end + 1) = ch;            % Python: self._bfr.append(char)
            end
        end

        function flush(obj)
            % FLUSH Write any pending <w:t> (run.py 303-307).
            %   Python: text = "".join(self._bfr); if text: self._r.add_t(text);
            %   self._bfr.clear().
            text = string(obj.bfr_);   % "".join -- reassembles surrogate pairs (H2)
            if strlength(text) > 0     % Python: `if text` (H4)
                obj.r_.add_t(text);
            end
            obj.bfr_ = '';             % self._bfr.clear()
        end
    end

    methods (Static)
        function append_to_run_from_text(r, text)
            % APPEND_TO_RUN_FROM_TEXT Append inner-content elements for `text` to
            %   `r` (run.py 275-279): appender = cls(r); appender.add_text(text).
            appender = mat2doc.oxml.text.RunContentAppender_(r);
            appender.add_text(string(text));
        end
    end
end
