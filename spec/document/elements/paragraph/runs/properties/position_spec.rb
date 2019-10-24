# frozen_string_literal: true

require 'spec_helper'

describe 'Position' do
  it 'position_negative' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/properties/position/position_negative.docx')
    expect(docx.elements.first.nonempty_runs[2].run_properties.position.value).to eq(OoxmlParser::OoxmlSize.new(-15, :half_point))
  end
end
