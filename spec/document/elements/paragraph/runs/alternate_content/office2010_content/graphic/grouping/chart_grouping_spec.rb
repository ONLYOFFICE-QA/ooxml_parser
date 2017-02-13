require 'spec_helper'

describe 'Chart Grouping' do
  it 'distance_from_text_in_emu' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/alternate_content/office2010_content/graphic/grouping/chart/chart_grouping.docx')
    expect(docx.elements.first.character_style_array[0].alternate_content.office2010_content.graphic.data).to be_a(OoxmlParser::ShapesGrouping)
  end
end
