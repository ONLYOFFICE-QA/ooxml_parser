require 'spec_helper'

describe 'page_properties' do
  describe 'page_borders' do
    it 'page borders offset' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_borders/page_border_offset.docx')
      expect(docx.page_properties.page_borders.offset_from).to eq(:page)
    end
  end

  describe 'page_margins' do
    it 'page margins parsing' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_margins/page_margins.docx')
      expect(docx.page_properties.margins).to eq(OoxmlParser::PageMargins.new(top: OoxmlParser::OoxmlSize.new(2, :centimeter),
                                                                              left: OoxmlParser::OoxmlSize.new(3, :centimeter),
                                                                              bottom: OoxmlParser::OoxmlSize.new(2, :centimeter),
                                                                              right: OoxmlParser::OoxmlSize.new(1.5, :centimeter),
                                                                              gutter: OoxmlParser::OoxmlSize.new(0, :centimeter),
                                                                              header: OoxmlParser::OoxmlSize.new(1.25, :centimeter),
                                                                              footer: OoxmlParser::OoxmlSize.new(1.25, :centimeter)))
    end
  end

  describe 'page_numbering' do
    it 'page_num_type_empty_format' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_numbering/page_num_type_empty_format.docx')
      expect(docx.page_properties.num_type).to be_nil
    end
  end
end
