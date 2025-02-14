export const qry_NewsPages = `
	<View>
		<Query>
			<Where>
				<And>
					<Eq>
						<FieldRef Name='FSObjType' />
						<Value Type='Integer'>0</Value>
					</Eq>
					<Eq>
						<FieldRef Name='IsNews' />
						<Value Type='Integer'>1</Value>
					</Eq>
				</And>
			</Where>
		</Query>
		<ViewFields>
			<FieldRef Name='Author' /><FieldRef Name='BannerImageUrl' /><FieldRef Name='Title' /><FieldRef Name='Created.' />
			<FieldRef Name='Created.FriendlyDisplay' /><FieldRef Name='Created_x0020_By' /><FieldRef Name='Editor' />
			<FieldRef Name='Description' /><FieldRef Name='EncodedAbsUrl' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' />
			<FieldRef Name='EncodedAbsUrl' /><FieldRef Name='Modified' /><FieldRef Name='Modified.' /><FieldRef Name='Modified.FriendlyDisplay' />
			<FieldRef Name='Modified_x0020_By' />
		</ViewFields>
	</View>
`;
export const qry_AllPages = `
	<View>
		<Query>
			<Where>
				<Eq>
					<FieldRef Name='FSObjType' />
					<Value Type='Integer'>0</Value>
				</Eq>
			</Where>
		</Query>
		<ViewFields>
			<FieldRef Name='Author' /><FieldRef Name='BannerImageUrl' /><FieldRef Name='Title' /><FieldRef Name='Created.' />
			<FieldRef Name='Created.FriendlyDisplay' /><FieldRef Name='Created_x0020_By' /><FieldRef Name='Editor' />
			<FieldRef Name='Description' /><FieldRef Name='EncodedAbsUrl' /><FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' />
			<FieldRef Name='EncodedAbsUrl' /><FieldRef Name='Modified' /><FieldRef Name='Modified.' /><FieldRef Name='Modified.FriendlyDisplay' />
			<FieldRef Name='Modified_x0020_By' />
		</ViewFields>
	</View>
`;
export const qry_AllNewsList = `
	<View>
		<ViewFields>
			<FieldRef Name='Author' /><FieldRef Name='CoverImage' /><FieldRef Name='Title' /><FieldRef Name='Created.' />
			<FieldRef Name='Created.FriendlyDisplay' /><FieldRef Name='Created_x0020_By' /><FieldRef Name='Editor' />
			<FieldRef Name='Description' /><FieldRef Name='NewsContent' /><FieldRef Name='IsActive' />
            <FieldRef Name='Modified' /><FieldRef Name='Modified.' /><FieldRef Name='Modified.FriendlyDisplay' />
			<FieldRef Name='Modified_x0020_By' />
		</ViewFields>
	</View>
`;
export const qry_ActiveNews = `
	<View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='IsActive' />
                    <Value Type='Integer'>1</Value>
                </Eq>
            </Where>
        </Query>
		<ViewFields>
			<FieldRef Name='Author' /><FieldRef Name='CoverImage' /><FieldRef Name='Title' /><FieldRef Name='Created.' />
			<FieldRef Name='Created.FriendlyDisplay' /><FieldRef Name='Created_x0020_By' /><FieldRef Name='Editor' />
			<FieldRef Name='Description' /><FieldRef Name='NewsContent' /><FieldRef Name='IsActive' />
            <FieldRef Name='Modified' /><FieldRef Name='Modified.' /><FieldRef Name='Modified.FriendlyDisplay' />
			<FieldRef Name='Modified_x0020_By' />
		</ViewFields>
	</View>
`;
export const qry_SitePages = `
	<View>
        <Query>
            <Where>
                <Eq>
                    <FieldRef Name='FSObjType' />
                    <Value Type='Integer'>0</Value>
                </Eq>
            </Where>
        </Query>
		<ViewFields>
			<FieldRef Name='Author' /><FieldRef Name='BannerImageUrl' /><FieldRef Name='BaseName' /><FieldRef Name='Created.' />
			<FieldRef Name='Created.FriendlyDisplay' /><FieldRef Name='Created_x0020_By' /><FieldRef Name='Editor' />
			<FieldRef Name='FileLeafRef' /><FieldRef Name='FileRef' /><FieldRef Name='Last_x0020_Modified.' />
            <FieldRef Name='Last_x0020_Modified' /><FieldRef Name='Modified.' /><FieldRef Name='Modified.FriendlyDisplay' />
			<FieldRef Name='Modified_x0020_By' /><FieldRef Name='PageLayoutType' /><FieldRef Name='Title' />
		</ViewFields>
	</View>
`;