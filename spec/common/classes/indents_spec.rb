# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::Indents do
  let(:indents) { described_class.new }

  it 'indents.to_s ' do
    expect(indents.to_s).to include('first line indent')
  end
end
