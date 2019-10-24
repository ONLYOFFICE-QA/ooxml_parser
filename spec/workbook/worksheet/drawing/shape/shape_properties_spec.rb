# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'Shape Line 2' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/shape_line.xlsx')
    expect(xlsx.worksheets.first.drawings.first.shape.properties.preset_geometry.name).to eq(:line)
  end

  it 'ShapesSpacing' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/drawing/shape/properties/shape_spacing.xlsx')
    expect(xlsx.worksheets[0].drawings[0].shape.text_body.paragraphs[0].properties.spacing).to eq(OoxmlParser::Spacing.new(OoxmlParser::OoxmlSize.new(4, :centimeter),
                                                                                                                           OoxmlParser::OoxmlSize.new(3, :centimeter),
                                                                                                                           OoxmlParser::OoxmlSize.new(5, :centimeter),
                                                                                                                           :exact))
  end
end
