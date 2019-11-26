# frozen_string_literal: true

require 'spec_helper'

describe 'Page Size' do
  describe 'Page Size Names' do
    it 'US Letter Portrait' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/us_letter_portrait.docx')
      expect(docx.page_properties.size.name).to eq('US Letter')
    end

    it 'US Letter Album' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/us_letter_album.docx')
      expect(docx.page_properties.size.name).to eq('US Letter')
    end

    it 'US Legal Album' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/us_legal.docx')
      expect(docx.page_properties.size.name).to eq('US Legal')
    end

    it 'A4' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/a4.docx')
      expect(docx.page_properties.size.name).to eq('A4')
    end

    it 'A5' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/a5.docx')
      expect(docx.page_properties.size.name).to eq('A5')
    end

    it 'B5' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/b5.docx')
      expect(docx.page_properties.size.name).to eq('B5')
    end

    it 'Envelope #10' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/envelope_10.docx')
      expect(docx.page_properties.size.name).to eq('Envelope #10')
    end

    it 'Envelope DL' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/envelope_dl.docx')
      expect(docx.page_properties.size.name).to eq('Envelope DL')
    end

    it 'Tabloid' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/tabloid.docx')
      expect(docx.page_properties.size.name).to eq('Tabloid')
    end

    it 'A3' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/a3.docx')
      expect(docx.page_properties.size.name).to eq('A3')
    end

    it 'Tabloid Oversize' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/tabloid_oversize.docx')
      expect(docx.page_properties.size.name).to eq('Tabloid Oversize')
    end

    it 'ROC 16K' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/roc_16k.docx')
      expect(docx.page_properties.size.name).to eq('ROC 16K')
    end

    it 'Envelope Choukei 3' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/envelope_choukei_3.docx')
      expect(docx.page_properties.size.name).to eq('Envelope Choukei 3')
    end

    it 'Super B/A3' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/super_b_a3.docx')
      expect(docx.page_properties.size.name).to eq('Super B/A3')
    end

    it 'Unrecognized Page Size' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/unrecognized_page_size.docx')
      expect(docx.page_properties.size.name).to include('Unknown page size')
    end
  end

  describe 'Other' do
    it 'page_size_custom.docx' do
      docx = OoxmlParser::DocxParser.parse_docx('spec/document/page_properties/page_size/page_size_custom.docx')
      expect(docx.page_properties.size.width).to eq(OoxmlParser::OoxmlSize.new(5, :centimeter))
      expect(docx.page_properties.size.height).to eq(OoxmlParser::OoxmlSize.new(29.7, :centimeter))
    end
  end
end
