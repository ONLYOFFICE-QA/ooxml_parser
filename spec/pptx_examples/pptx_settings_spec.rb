require 'rspec'
require 'ooxml_parser'

describe '#configure' do
  it 'Default Unit of measurement is centimeter' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(16.5)
  end

  it 'Set unit of measurement as points' do
    OoxmlParser.configure do |config|
      config.units = :points
    end
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(467.717)
  end

  it 'Set unit of measurement as dxa' do
    OoxmlParser.configure do |config|
      config.units = :dxa
    end
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(5_940_000)
  end

  it 'Set unit of measurement as inches' do
    OoxmlParser.configure do |config|
      config.units = :inches
    end
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(4125)
  end

  it 'Set unit of measurement as emu' do
    OoxmlParser.configure do |config|
      config.units = :emu
    end
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/pptx_examples/settings/measurement.pptx')
    expect(pptx.slides[0].elements.last.properties.transform.extents.x).to eq(3_771_900_000)
  end

  after do
    OoxmlParser.reset
  end
end
