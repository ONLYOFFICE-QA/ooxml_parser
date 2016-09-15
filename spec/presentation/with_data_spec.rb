require 'spec_helper'

describe 'Presentation#with_data?' do
  it 'no_data' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/no_data.pptx')
    expect(pptx).not_to be_with_data
  end

  it 'second_slide' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/second_slide.pptx')
    expect(pptx).to be_with_data
  end

  it 'slide_with_text' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/slide_with_text.pptx')
    expect(pptx).to be_with_data
  end

  it 'shape' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/shape.pptx')
    expect(pptx).to be_with_data
  end

  it 'table' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/table.pptx')
    expect(pptx).to be_with_data
  end

  it 'chart' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/chart.pptx')
    expect(pptx).to be_with_data
  end

  it 'docx_picture' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/docx_picture.pptx')
    expect(pptx).to be_with_data
  end

  it 'blip_fill' do
    pptx = OoxmlParser::Parser.parse('spec/presentation/with_data/blip_fill.pptx')
    pptx.with_data?
    expect(pptx).to be_with_data
  end
end
