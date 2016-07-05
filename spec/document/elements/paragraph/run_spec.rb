require 'spec_helper'

describe OoxmlParser::ParagraphRun do
  it 'caps_characters_on' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/caps_characters_on.docx')
    expect(docx.elements.first.character_style_array.first.caps).to eq(:caps)
  end

  describe 'footnote' do
    it 'footnote_reference' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/footnote_reference.docx')
      expect(docx.elements[91].character_style_array[23].footnote.id).to eq(1)
    end
  end
end
