require 'spec_helper'

describe '#configure' do
  it 'Default Unit of measurement is centimeter' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(16.5, :centimeter))
  end

  it 'Set unit of measurement as points' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(467.72, :point))
  end

  it 'Set unit of measurement as dxa' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(9354.33, :dxa))
  end

  it 'Set unit of measurement as inches' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(6.5, :inch))
  end

  it 'Set unit of measurement as emu' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(OoxmlParser::OoxmlSize.new(5_940_000, :emu))
  end
end
