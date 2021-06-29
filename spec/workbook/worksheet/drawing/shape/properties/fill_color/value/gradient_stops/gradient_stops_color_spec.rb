# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'gradient_fill.xlsx' do
    xlsx = OoxmlParser::XlsxParser.parse_xlsx('spec/workbook/worksheet/'\
                                              'drawing/shape/properties/fill_color/'\
                                              'value/gradient_stops/gradient_stops_color'\
                                              '/gradient_stops_color_preset.xlsx')
    expect(xlsx.worksheets[0].drawings.first
               .shape.properties.fill_color
               .value.gradient_stops
               .first.color.value).to eq('lightYellow')
  end
end
