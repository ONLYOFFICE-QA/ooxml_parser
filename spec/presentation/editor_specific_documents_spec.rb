# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  describe 'libreoffice' do
    it 'simple_text' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/editor_specific_documents/libreoffice/simple_text.pptx')
      drawing = pptx.slides[0].elements.last
      expect(drawing.text_body.paragraphs.first.runs.first.text).to eq('This is a test')
    end
  end

  describe 'ms_office_2013' do
    it 'simple_text' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/editor_specific_documents/ms_office_2013/simple_text.pptx')
      drawing = pptx.slides[0].elements.last
      expect(drawing.text_body.paragraphs.first.runs.first.text).to eq('This is a test')
    end
  end
end
