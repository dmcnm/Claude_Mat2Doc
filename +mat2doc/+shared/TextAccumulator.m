classdef TextAccumulator < handle
% TEXTACCUMULATOR Accepts str fragments and joins them together on .pop().
%
%   Accepts `str` fragments and joins them together, in order, on `.pop()`.
%   Handy when text in a stream is broken up arbitrarily and you want to join
%   it back together within certain bounds. The optional `separator` argument
%   determines how the text fragments are punctuated, defaulting to the empty
%   string.
%
%   NEW FOR MAT2DOC (audit_P2-1_proxy_tier.md): no python-pptx counterpart.
%   Used later by the text-extraction tier; ported faithfully here as part of
%   the shared.py tier.
%
%   REFERENCE SEMANTICS: a handle class -- push/pop mutate a shared internal
%   buffer, mirroring Python's mutable TextAccumulator instance.
%
%   GENERATOR PORT (H9): Python's `pop()` is a GENERATOR that yields ZERO or
%   ONE str (used as `yield from accum.pop()` so an empty accumulator produces
%   nothing rather than an empty string). Per design.md section 2 (generators
%   -> precomputed arrays), MATLAB `pop()` returns a (1,:) string of length 0
%   (when the buffer is empty) or length 1 (the joined fragments). A consumer
%   ports `yield from accum.pop()` as iterating the returned array
%   (`for t = accum.pop()`), which yields nothing for a 1x0 result -- exactly
%   the Python behavior.
%
%   UNDERSCORE ROTATION (design.md section 2): the private `_separator` /
%   `_texts` attributes rotate the leading underscore -> separator_ / texts_.
%
%   Example:
%       acc = mat2doc.shared.TextAccumulator();
%       acc.push("foo");
%       acc.push("bar");
%       t = acc.pop();      % "foobar" (1x1 string); buffer now empty
%       e = acc.pop();      % 1x0 string (nothing accumulated)
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::TextAccumulator

    properties (Access = private)
        separator_ (1,1) string = ""    % _separator: the join delimiter
        texts_ (1,:) string             % _texts: accumulated fragments (insertion order)
    end

    methods
        function obj = TextAccumulator(separator)
            % TEXTACCUMULATOR Construct with an optional join separator
            %   (shared.py 364-366): self._separator = separator; self._texts = [].
            %
            %   Inputs:  separator - (optional) (1,1) string join delimiter.
            %                        Default "" (concatenation).
            %   Outputs: obj       - a scalar TextAccumulator handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::TextAccumulator.__init__
            arguments
                separator (1,1) string = ""
            end
            obj.separator_ = separator;
            obj.texts_ = strings(1, 0);     % empty buffer (Python [])
        end

        function push(obj, text)
            % PUSH Add a text fragment to the accumulator (shared.py 368-370):
            %   self._texts.append(text).
            %
            %   Inputs:  text - (1,1) string fragment to append.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::TextAccumulator.push
            arguments
                obj
                text (1,1) string
            end
            obj.texts_(end + 1) = text;     % append (insertion order preserved)
        end

        function out = pop(obj)
            % POP Generate zero-or-one str from those accumulated (shared.py 372-382).
            %
            %   Python: if not self._texts: return (yields nothing); else join the
            %   fragments with the separator, clear the buffer, and yield the one
            %   string. Ported (H9) as a (1,:) string of length 0 or 1.
            %
            %   Ported from python-docx v1.2.0: src/docx/shared.py::TextAccumulator.pop
            if isempty(obj.texts_)          % `if not self._texts` (H4: empty list falsy)
                out = strings(1, 0);        % yield nothing
                return
            end
            % self._separator.join(self._texts)
            out = join(obj.texts_, obj.separator_);
            obj.texts_ = strings(1, 0);     % self._texts.clear()
        end
    end
end
