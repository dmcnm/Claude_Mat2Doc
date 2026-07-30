classdef Text_ < handle
% TEXT_ Proxy object wrapping a `<w:t>` element (Python `_Text`).
%
%   The small wrapper returned by Run.add_text. In python-docx v1.2.0 `_Text`
%   is a bare `object` subclass whose ONLY state is the wrapped `w:t` element
%   -- it exposes NO `text` property or any other member (verified against the
%   v1.2.0 source, run.py 252-257, and the runtime oracle: `hasattr(t,"text")`
%   is False). It is ported faithfully as an inert handle holding `t_`; NO
%   `text` accessor is added (design.md section 7: no features beyond the
%   original). REFERENCE SEMANTICS: a handle class (an API proxy over an
%   element); Python `_Text` instances are reference objects.
%
%   UNDERSCORE ROTATION (design.md section 2): the module-private class `_Text`
%   rotates the leading underscore to the END -> Text_; the private `_t`
%   attribute -> t_.
%
%   Python `super(_Text, self).__init__()` is just `object.__init__()` and has
%   no ported effect.
%
%   Example:
%       r = mat2doc.oxml.OxmlElement("w:r");
%       run = mat2doc.text.Run(r, someParagraph);
%       t = run.add_text(" hi ");    % a mat2doc.text.Text_ over the new <w:t>
%
%   Ported from python-docx v1.2.0: src/docx/text/run.py::_Text

    properties (Access = private)
        t_          % _t (run.py 257): the wrapped <w:t> (a mat2doc.oxml.text.CT_Text)
    end

    methods
        function obj = Text_(t_elm)
            % TEXT_ Wrap a `<w:t>` element (run.py 255-257).
            %
            %   Inputs:  t_elm - a mat2doc.oxml.text.CT_Text (the `w:t` element).
            %   Outputs: obj   - a scalar Text_ handle.
            %
            %   Ported from python-docx v1.2.0: src/docx/text/run.py::_Text.__init__
            obj.t_ = t_elm;             % Python: self._t = t_elm
        end
    end
end
