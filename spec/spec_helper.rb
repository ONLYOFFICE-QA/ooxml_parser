# frozen_string_literal: true

require 'simplecov'
SimpleCov.start do
  add_filter('/spec/')
end

require 'ooxml_parser'

RSpec.configure do |config|
  config.after do |_example|
    GC.start
  end
end
