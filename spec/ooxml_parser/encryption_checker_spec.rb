# frozen_string_literal: true

require 'spec_helper'

describe OoxmlParser::EncryptionChecker do
  it 'encrypted?(path, ignore_system: true) always return false' do
    expect(described_class.new(__FILE__, ignore_system: true)).not_to be_encrypted
  end
end
