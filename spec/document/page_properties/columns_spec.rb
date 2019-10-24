# frozen_string_literal: true

require 'spec_helper'

describe 'Columns' do
  it 'two_columns' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/two_columns.docx')
    expect(docx.page_properties.columns.count).to eq(2)
  end

  it 'ten_columns' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/ten_columns.docx')
    expect(docx.page_properties.columns.count).to eq(10)
  end

  it 'left_column' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/left_column.docx')
    expect(docx.page_properties.columns[0].width).to eq(OoxmlParser::OoxmlSize.new(4.68, :centimeter))
  end

  it 'right_column' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/right_column.docx')
    expect(docx.page_properties.columns[0].width).to eq(OoxmlParser::OoxmlSize.new(10.63, :centimeter))
  end

  it 'several_types_of_columns' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/several_types_of_columns.docx')
    expect(docx.elements[2].sector_properties.columns.count).to eq(2)
    expect(docx.elements[5].sector_properties.columns.count).to eq(3)
  end

  it 'columns_space' do
    docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/columns_space.docx')
    expect(docx.page_properties.columns.space).to eq(OoxmlParser::OoxmlSize.new(0.5, :inch))
  end

  describe 'equal_columns' do
    it 'equal_columns' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/equal_columns/equal_columns_undefined.docx')
      expect(docx.page_properties.columns.equal_width).to be_nil
    end

    it 'non_equal_columns' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/equal_columns/non_equal_columns.docx')
      expect(docx.page_properties.columns).not_to be_equal_width
    end

    it 'equal_columns_undefined' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/equal_columns/equal_columns.docx')
      expect(docx.page_properties.columns).to be_equal_width
    end
  end
end
