# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::ParagraphRun do
  it 'caps_characters_on' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/caps_characters_on.docx')
    expect(docx.elements.first.character_style_array.first.caps).to eq(:caps)
  end

  it 'font properties' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/run_font_properties.docx')
    expect(docx.elements.first.nonempty_runs[0].text).to eq('This is a paragraph with the text color, font family and font size set using the text style. ')
    expect(docx.elements.first.nonempty_runs[0].font).to eq('Calibri Light')
    expect(docx.elements.first.nonempty_runs[0].font_color).to eq(OoxmlParser::Color.new(38, 38, 38))
    expect(docx.elements.first.nonempty_runs[1].text).to eq('We do not change the style of the paragraph itself. ')
    expect(docx.elements.first.nonempty_runs[1].font_color).to eq(OoxmlParser::Color.new(38, 38, 38))
    expect(docx.elements.first.nonempty_runs[1].font).to eq('Calibri Light')
    expect(docx.elements.first.nonempty_runs[2].text).to eq('Only document-wide text styles are applied.')
    expect(docx.elements.first.nonempty_runs[2].font_color).to eq(OoxmlParser::Color.new(38, 38, 38))
    expect(docx.elements.first.nonempty_runs[2].font).to eq('Calibri Light')
  end
end
