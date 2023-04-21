# frozen_string_literal: true

require 'spec_helper'
describe OoxmlParser::Note do
  describe '#note_base_xpath' do
    it 'returns the XPath for footer notes when type includes "footer"' do
      note = described_class.new(type: 'footer')
      expect(note.note_base_xpath).to eq('//w:ftr')
    end

    it 'returns the XPath for header notes when type includes "header"' do
      note = described_class.new(type: 'header')
      expect(note.note_base_xpath).to eq('//w:hdr')
    end

    it 'raises a NameError when type is unknown' do
      note = described_class.new(type: 'unknown')
      expect { note.note_base_xpath }.to raise_error(NameError, 'Unknown note type: unknown')
    end
  end
end
