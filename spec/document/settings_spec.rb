# frozen_string_literal: true

require 'spec_helper'

describe 'DocumentSettings' do
  it 'no_settings_xml' do
    docx = OoxmlParser::Parser.parse('spec/document/settings/no_setting_xml.docx')
    expect(docx.settings).to be_nil
  end
end
