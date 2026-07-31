classdef Settings < mat2doc.shared.ElementProxy
% SETTINGS Provides access to document-level settings for a document.
%
%   Accessed using the Document.settings property (document.py 211-214). A thin
%   ElementProxy over the `<w:settings>` root element (a CT_Settings) of the
%   settings part; reference semantics (handle) and H5 element-identity eq/ne
%   are inherited from ElementProxy unchanged.
%
%   The single ported member odd_and_even_pages_header_footer (read/write)
%   delegates straight to the CT_Settings element's evenAndOddHeaders_val
%   accessor -- no local state. Ported as a Dependent property with get./set.
%   (house convention for a read/write proxy @property, e.g. Font.bold), so
%   Python `st.odd_and_even_pages_header_footer = True` mirrors as MATLAB
%   `st.odd_and_even_pages_header_footer = true`. `element` is inherited from
%   ElementProxy (a zero-arg method).
%
%   H3 (tri-state): odd_and_even_pages_header_footer is a plain bool at this
%   layer (True/False), backed by the CT_Settings None/True/False tri-state.
%
%   Example:
%       d  = mat2doc.Document();
%       st = d.settings;
%       st.odd_and_even_pages_header_footer            % false by default
%       st.odd_and_even_pages_header_footer = true;    % -> <w:evenAndOddHeaders/>
%
%   Ported from python-docx v1.2.0: src/docx/settings.py::Settings

    properties (Access = private)
        settings_        % _settings: element cast to CT_Settings (settings.py 23)
    end

    properties (Dependent)
        odd_and_even_pages_header_footer   % bool; distinct odd/even headers & footers
    end

    methods
        function obj = Settings(element, parent)
            % SETTINGS Wrap a `w:settings` element (settings.py 21-23).
            %
            %   Inputs:  element - a mat2doc.oxml.settings.CT_Settings.
            %            parent  - (optional) a ProvidesXmlPart. Default [] (None).
            %   Outputs: obj     - a scalar Settings handle.
            %
            %   Python: super().__init__(element, parent);
            %           self._settings = cast("CT_Settings", element).
            %
            %   Ported from python-docx v1.2.0: src/docx/settings.py::Settings.__init__
            arguments
                element
                parent = []     % None sentinel (H3): parent defaults to None
            end
            obj@mat2doc.shared.ElementProxy(element, parent);   % Python: super().__init__(element, parent)
            obj.settings_ = element;   % Python: self._settings = cast("CT_Settings", element)
        end

        function value = get.odd_and_even_pages_header_footer(obj)
            % Python (settings.py 25-31, @property):
            %   return self._settings.evenAndOddHeaders_val
            value = obj.settings_.evenAndOddHeaders_val;
        end

        function set.odd_and_even_pages_header_footer(obj, value)
            % Python (settings.py 33-35, @property.setter):
            %   self._settings.evenAndOddHeaders_val = value
            obj.settings_.evenAndOddHeaders_val = value;
        end
    end
end
