# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocxBlip do
  let(:docx_blip) { described_class.new }

  it '#to_str' do
    string = 'foo'
    docx_blip.path_to_media_file = string
    expect(docx_blip.to_str).to eq(string)
  end
end
