# frozen_string_literal: true

require 'spec_helper'

describe 'nonempty_runs' do
  it 'endnote_nonempty_runs' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/nonempty_runs/endnote_nonempty_runs.docx')
    expect(docx.elements.first.nonempty_runs.length).to eq(2)
  end

  it 'footnote_nonempty_runs' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/nonempty_runs/footnote_nonempty_runs.docx')
    expect(docx.elements.first.nonempty_runs.length).to eq(4)
  end

  it 'nonempty_runs_with_shape' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/nonempty_runs/nonempty_runs_with_shape.docx')
    expect(docx.elements.last.nonempty_runs.length).to eq(4)
  end

  it 'sdt_as_paragraph_element' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/nonempty_runs/sdt_as_paragraph_element.docx')
    expect(docx.elements.first.nonempty_runs.first).to be_a(OoxmlParser::StructuredDocumentTag)
  end
end
