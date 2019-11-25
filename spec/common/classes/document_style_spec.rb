# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::DocumentStyle do
  let(:style) do
    OoxmlParser::DocumentStyle.new
  end

  it 'DocumentStyle#to_s is correct' do
    expect(style.to_s).to include('Table style properties list')
  end

  it 'DocumentStyle#inspect is correct' do
    expect(style.inspect).to include('Table style properties list')
  end
end
