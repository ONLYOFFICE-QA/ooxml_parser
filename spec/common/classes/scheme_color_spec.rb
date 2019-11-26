# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::SchemeColor do
  let(:scheme) { described_class.new }

  it 'OoxmlParser::SchemeColor#to_s is correct for new' do
    expect(scheme.to_s).to eq('0')
  end
end
