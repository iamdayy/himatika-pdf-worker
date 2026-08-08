import ast

with open('api/index.py', 'r') as f:
    source = f.read()
    lines = source.splitlines()

tree = ast.parse(source)

functions = []
for node in tree.body:
    if isinstance(node, ast.FunctionDef):
        # find decorators
        decorators = [ast.unparse(d) for d in node.decorator_list]
        is_route = any('app.route' in d for d in decorators)
        functions.append({
            'name': node.name,
            'start': node.lineno - 1 - len(node.decorator_list), # approximation
            'end': node.end_lineno,
            'is_route': is_route
        })

for f in functions:
    if f['is_route']:
        print(f"{f['name']}: {f['start']} - {f['end']}")
