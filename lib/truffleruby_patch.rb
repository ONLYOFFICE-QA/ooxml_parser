# frozen_string_literal: true

# Monkey patch around
# https://github.com/oracle/truffleruby/commit/cc76155bf509587e1b2954e9de77c558c7c857f4
# Remove it as soon as TruffleRuby with fix is released
if defined?(Truffle)
  # MonkeyPatch File stdlib class
  class File < IO
    # Monkey patch constant missing on TruffleRuby
    SHARE_DELETE = 0
  end
end
