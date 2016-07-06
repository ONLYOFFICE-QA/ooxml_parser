require 'spec_helper'

describe 'My behaviour' do
  it 'Character Spacing 1cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_1_cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(1)
  end

  it 'Character Spacing 5cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_5_cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(5)
  end

  it 'Character Spacing 0.1cm' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_0.1cm.pptx')
    paragraph = pptx.slides.first.elements.first.text_body.paragraphs.first
    expect(paragraph.characters.first.properties.space).to eq(0.1)
  end

  it 'Spacing' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0.35, 1.15, :multiple))
  end

  it 'DefaultSpace' do
    pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/default_space.pptx')
    expect(pptx.slides.first.elements.last.text_body.paragraphs.first.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 1, :multiple))
  end

  describe 'Spacing Rule' do
    it 'spacing_line_rule_exact.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_line_rule_exact.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(1.84, 1.92, 0.8, :exact))
    end

    it 'spacing_line_rule_exact_0.35.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_line_rule_exact_0.35.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 0.35, :exact))
    end

    it 'spacing_line_rule_exact_1.76.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_line_rule_exact_1.76.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(0, 0, 1.76, :exact))
    end

    it 'spacing_line_rule_multiple.pptx' do
      pptx = OoxmlParser::PptxParser.parse_pptx('spec/presentation/slide/shape/text_body/paragraphs/runs/properties/space/spacing_line_rule_multiple.pptx')
      paragraph = pptx.slides.first.elements.last.text_body.paragraphs.first
      expect(paragraph.properties.spacing).to eq(OoxmlParser::Spacing.new(1.84, 1.92, 1, :multiple))
    end
  end
end
