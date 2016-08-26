require 'spec_helper'

describe 'Shade' do
  it 'should do something' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/elements/paragraph/runs/properties/shade/shade_clear.docx')
    expect(docx.elements.first.nonempty_runs[1].run_properties.shade.value).to eq(:clear)
    expect(docx.elements.first.nonempty_runs[1].run_properties.shade.color).to eq(:auto)
    expect(docx.elements.first.nonempty_runs[1].run_properties.shade.fill).to eq(OoxmlParser::Color.new(0, 255, 0))
  end
end
