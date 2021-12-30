# frozen_string_literal: true

require 'spec_helper'

describe 'Password protected documents' do
  it 'Password protected docx' do
    expect do
      docx = OoxmlParser::DocxParser.parse_docx('spec/common/password_protected/password_protected.docx')
      expect(docx).to be_nil
    end.to output(/is encrypted/).to_stderr
  end

  it 'Checking for password protection file with space in name' do
    expect do
      docx = OoxmlParser::DocxParser.parse_docx('spec/common/password_protected/password protected space in name.docx')
      expect(docx).to be_nil
    end.to output(/is encrypted/).to_stderr
  end

  it 'Password protected xlsx' do
    expect do
      docx = OoxmlParser::DocxParser.parse_docx('spec/common/password_protected/password_protected.xlsx')
      expect(docx).to be_nil
    end.to output(/is encrypted/).to_stderr
  end

  it 'Password protected pptx' do
    expect do
      docx = OoxmlParser::DocxParser.parse_docx('spec/common/password_protected/password_protected.pptx')
      expect(docx).to be_nil
    end.to output(/is encrypted/).to_stderr
  end
end
