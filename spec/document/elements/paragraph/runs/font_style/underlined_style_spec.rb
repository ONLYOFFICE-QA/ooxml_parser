# frozen_string_literal: true

require 'spec_helper'

describe 'Underlined' do
  it 'underlined_single' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_style/underlined/underlined_single.docx')
    expect(docx.elements[0].nonempty_runs.first.font_style.underlined).to eq(:single)
  end

  it 'underline is a symbol' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/font_style/underlined/underline_single_symbol.docx')
    expect(docx.notes.first.elements[1].nonempty_runs.first.text).to eq('Underline')
    expect(docx.notes.first.elements[1].nonempty_runs.first.font_style.underlined.style).to eq(:single)
  end
end
