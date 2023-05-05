# frozen_string_literal: true

require 'spec_helper'

describe 'chart', '#data_point_marker' do
  file = 'spec/document/elements/paragraph/runs/' \
         'drawing/graphic/chart/chart_with_style.docx'
  docx = OoxmlParser::Parser.parse(file)
  layout = docx.elements[0].nonempty_runs.first.drawing.graphic.data.style.data_point_marker_layout

  it 'Layout has symbol' do
    expect(layout.symbol).to eq(:circle)
  end

  it 'Layout has size' do
    expect(layout.size).to eq(6)
  end
end
