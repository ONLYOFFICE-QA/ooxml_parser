require 'spec_helper'

describe 'My behaviour' do
  it 'Check theme name Стандартная' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/theme/theme_standart.pptx')
    expect(pptx.theme.name).to eq('Стандартная')
  end

  it 'NoThemeName' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/theme/no_theme_name.pptx')
    expect(pptx.theme.name).to eq('')
  end

  it 'ClassicThemeName' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/theme/classic_theme.pptx')
    expect(pptx.theme.name).to eq('Classic')
  end
end
