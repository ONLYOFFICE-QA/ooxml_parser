# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Spacing do
  files_dir = 'spec/document/elements/paragraph/spacing'

  it 'ShapeParagraphSpacing_1' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/shape_paragraph_spacing_1.docx")
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing)
      .to eq(described_class.new(0.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagraphSpacing_2' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/shape_paragraph_spacing_2.docx")
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing)
      .to eq(described_class.new(10.0, 0.35, 1.0, :auto))
  end

  it 'ShapeParagrpaphSpacing_3' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/shape_paragraph_spacing_3.docx")
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing)
      .to eq(described_class.new(0, 0.35, 1.5, :auto))
  end

  it 'ShapeDefaultSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/space_default_spacing.docx")
    expect(docx.element_by_description(location: :canvas, type: :shape).first.spacing)
      .to eq(described_class.new(0, 0.35, 1.15, :auto))
  end

  it 'TableLineSpacing' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/table_line_spacing.docx")
    expect(docx.element_by_description(location: :canvas, type: :table)[0].spacing.line).to eq(1.0)
  end

  it 'Line spacing if order of attribute changed' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/line_spacing_another_order.docx")
    expect(docx.elements.first.spacing.line).to eq(3.0)
  end

  it 'Parse default spacing' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/default_spacing.docx")
    expect(docx.elements.first.spacing).to eq(described_class.new(0, 0.35, 1.15, :auto))
  end

  it 'Parsing line rule' do
    docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/line_rule.docx")
    expect(docx.elements.first.spacing.line_rule).to eq(:auto)
  end

  describe 'contextual spacing' do
    it 'contextual_spacing false' do
      docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/contextual_spacing_false.docx")
      expect(docx.elements.first.contextual_spacing).to be_falsey
    end

    it 'contextual_spacing true' do
      docx = OoxmlParser::DocxParser.parse_docx("#{files_dir}/contextual_spacing_true.docx")
      expect(docx.document_style_by_name('Title').paragraph_properties.contextual_spacing).to be_truthy
    end
  end
end
