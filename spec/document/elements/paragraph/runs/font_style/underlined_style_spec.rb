# frozen_string_literal: true

require 'spec_helper'

describe 'Underlined' do
  it 'underlined_single' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/font_style/underlined/underlined_single.docx')
    expect(docx.elements[0].nonempty_runs.first.font_style.underlined).to eq(:single)
  end
end
