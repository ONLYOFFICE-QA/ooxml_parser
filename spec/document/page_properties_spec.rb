require 'spec_helper'

describe 'page_properties' do
  describe 'columns' do
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
      expect(docx.page_properties.columns[0].width).to eq(0.01)
    end

    it 'right_column' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/right_column.docx')
      expect(docx.page_properties.columns[0].width).to eq(0.02)
    end

    it 'several_types_of_columns' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/columns/several_types_of_columns.docx')
      expect(docx.elements[2].sector_properties.columns.count).to eq(2)
      expect(docx.elements[5].sector_properties.columns.count).to eq(3)
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

  describe 'page_borders' do
    it 'page borders offset' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_borders/page_border_offset.docx')
      expect(docx.page_properties.page_borders.offset_from).to eq(:page)
    end
  end

  describe 'page_margins' do
    it 'page margins parsing' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_margins/page_margins.docx')
      expect(docx.page_properties.margins).to eq(OoxmlParser::PageMargins.new(top: 2, left: 3, bottom: 2, right: 1.5, gutter: 0, header: 1.25, footer: 1.25))
    end
  end

  describe 'page_numbering' do
    it 'page_num_type_empty_format' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_numbering/page_num_type_empty_format.docx')
      expect(docx.page_properties.num_type).to be_nil
    end
  end
end
