require 'spec_helper'

describe 'Drawing Properties Shape Size' do
  let(:docx) { OoxmlParser::Parser.parse('spec/document/elements/paragraph/runs/drawing/properties/shape_size/shape_size_flip.docx') }
  let(:shape_size) { docx.elements.first.nonempty_runs[0].alternate_content.office2010_content.graphic.data.properties.shape_size }

  it 'shape_size_flip_horizontally is boolean and true' do
    expect(shape_size.flip_horizontal).to be_truthy
  end

  it 'shape_size_flip_vertically is boolean and false' do
    expect(shape_size.flip_vertical).to be_falsey
  end
end
