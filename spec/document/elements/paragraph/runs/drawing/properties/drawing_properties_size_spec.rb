require 'spec_helper'

describe 'Drawing Properties Size' do
  it 'size_relative' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/drawing/properties/size/size_relative.docx')
    shape_properties = docx.elements.first.nonempty_runs.first.alternate_content.office2010_content.properties
    expect(shape_properties.size_relative_horizontal.relative_from).to eq(:right_margin)
    expect(shape_properties.size_relative_horizontal.width.value).to eq(OoxmlParser::OoxmlSize.new(50, :percent))
    expect(shape_properties.size_relative_vertical.relative_from).to eq(:page)
    expect(shape_properties.size_relative_vertical.height.value).to eq(OoxmlParser::OoxmlSize.new(50, :percent))
  end
end
