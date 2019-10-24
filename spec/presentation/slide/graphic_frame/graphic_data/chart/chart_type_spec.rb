# frozen_string_literal: true

require 'spec_helper'

describe 'My behaviour' do
  it 'column_charts.pptx' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/graphic_frame/graphic_data/chart/type/column_charts.pptx')
    elements = pptx.slides.first.elements
    expect(elements[0].graphic_data.first.type).to eq(:column)
    expect(elements[0].graphic_data.first.grouping).to eq(:clustered)
    expect(elements[1].graphic_data.first.type).to eq(:column)
    expect(elements[1].graphic_data.first.grouping).to eq(:stacked)
    expect(elements[2].graphic_data.first.type).to eq(:column)
    expect(elements[2].graphic_data.first.grouping).to eq(:percentStacked)
  end
end
