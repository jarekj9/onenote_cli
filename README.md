Project uses microsoft graph ( https://developer.microsoft.com/en-us/graph/graph-explorer )
It requires, that account has specific permissions to onenote ( https://docs.microsoft.com/en-us/graph/permissions-reference )

To download note book use:
```./onenote.py -u 'username@outlook.com'```

To browse chech help with '-h' flag.

Example - show page with title 'cron' in notebook section 'LINUX':

```./onenote.py -s LINUX -t cron```
