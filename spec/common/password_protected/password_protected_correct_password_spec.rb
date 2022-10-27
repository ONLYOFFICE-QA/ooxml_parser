# frozen_string_literal: true

require 'spec_helper'

password = '111'

describe 'Password protected documents with correct password' do
  before do
    pending('https://github.com/woodbusy/ooxml_decrypt/issues/6') if RUBY_PLATFORM == 'java'
  end

  it 'Password protected docx' do
    docx = OoxmlParser::Parser.parse('spec/common/password_protected/files/password_protected.docx', password: password)
    expect(docx).not_to be_nil
  end

  it 'Password protected xlsx' do
    xlsx = OoxmlParser::Parser.parse('spec/common/password_protected/files/password_protected.xlsx', password: password)
    expect(xlsx).not_to be_nil
  end

  it 'Password protected pptx' do
    pptx = OoxmlParser::Parser.parse('spec/common/password_protected/files/password_protected.pptx', password: password)
    expect(pptx).not_to be_nil
  end
end
