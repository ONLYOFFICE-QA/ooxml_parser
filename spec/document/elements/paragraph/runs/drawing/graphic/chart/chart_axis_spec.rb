require 'spec_helper'

describe 'Chart axises' do
  it 'Check Image with empty link' do
    docx = OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/drawing/graphic/chart/axis/root_element_of_chart_axis.docx')
    expect(docx).to be_with_data
  end
end
