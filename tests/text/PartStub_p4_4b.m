classdef PartStub_p4_4b < handle
% PARTSTUB_P4_4B  Minimal `part` provider for the P4-4b Run tests.
%   A local twin of the Gate-3 s0013_PartStub (validation-folder, NOT on the
%   toolbox path): a ProvidesStoryPart structurally -- exposes only a
%   zero-argument part() accessor returning a sentinel, so StoryChild.part
%   delegation (Run(...).part -> parent_.part()) is observable without pulling
%   in a real DocumentPart. NOT a matlab.unittest.TestCase, so runtests
%   contributes zero cases from this file.
    methods
        function p = part(~)
            p = "STUBPART";
        end
    end
end
