# frozen_string_literal: true

require 'spec_helper'

describe 'Spacing' do
  it 'spacing_sizing' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/properties/spacing/spacing_sizing.docx')
    expect(docx.elements.first.nonempty_runs[1].run_properties.spacing.value).to eq(OoxmlParser::OoxmlSize.new(80, :twip))
  end
end
