# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ParagraphRun do
  it 'caps_characters_on' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/caps_characters_on.docx')
    expect(docx.elements.first.character_style_array.first.caps).to eq(:caps)
  end
end
