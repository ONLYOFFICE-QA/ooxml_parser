# frozen_string_literal: true

require 'simplecov'
SimpleCov.start do
  add_filter('/spec/')
  # Remove as soon as
  # https://github.com/oracle/truffleruby/commit/cc76155bf509587e1b2954e9de77c558c7c857f4
  # is released
  add_filter('lib/truffleruby_patch.rb')
end

require 'ooxml_parser'
