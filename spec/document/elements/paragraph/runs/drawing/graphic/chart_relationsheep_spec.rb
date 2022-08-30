# frozen_string_literal: true

require 'spec_helper'

describe 'chart', '#relationship' do
  file = 'spec/document/elements/paragraph/runs/' \
         'drawing/graphic/chart/chart_with_style.docx'

  it 'Chart has nonempty relationship object' do
    docx = OoxmlParser::Parser.parse(file)
    chart = docx.elements[0].nonempty_runs.first.drawing.graphic.data
    expect(chart.relationships.relationship.length).to eq(3)
  end
end
