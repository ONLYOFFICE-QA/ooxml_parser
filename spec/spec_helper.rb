# frozen_string_literal: true

# Until https://github.com/oracle/truffleruby/issues/2813 is released
unless RUBY_ENGINE == 'truffleruby'
  require 'simplecov'
  SimpleCov.start do
    add_filter('/spec/')
  end
end

require 'ooxml_parser'
