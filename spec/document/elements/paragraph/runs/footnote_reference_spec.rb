# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'footnote_reference_id' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/footnote_references/footnote_reference_id.docx')
    expect(docx.elements[91].character_style_array[23].footnote.id).to eq(1)
  end

  it 'footnote_reference_paragraph' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/footnote_references/footnote_reference_paragraph.docx')
    expect(docx.elements[91].character_style_array[23].footnote.elements.first).to be_a(OoxmlParser::DocxParagraph)
  end
end
