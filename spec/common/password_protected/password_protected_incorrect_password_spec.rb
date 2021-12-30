# frozen_string_literal: true

require 'spec_helper'

password = '1'

describe 'Password protected documents with incorrect password' do
  it 'Password protected docx' do
    path = 'spec/common/password_protected/files/password_protected.docx'
    expect { OoxmlParser::Parser.parse(path, password: password) }.to raise_error(/Zip end of central directory signature not found/)
  end

  it 'Password protected xlsx' do
    path = 'spec/common/password_protected/files/password_protected.xlsx'
    expect { OoxmlParser::Parser.parse(path, password: password) }.to raise_error(/Zip end of central directory signature not found/)
  end

  it 'Password protected pptx' do
    path = 'spec/common/password_protected/files/password_protected.pptx'
    expect { OoxmlParser::Parser.parse(path, password: password) }.to raise_error(/Zip end of central directory signature not found/)
  end
end
